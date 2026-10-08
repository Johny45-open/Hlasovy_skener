import glob
import logging
import threading
import os
from dataclasses import dataclass
from typing import Optional

import numpy as np
import pytesseract
import easyocr
from PIL import Image
from PyQt6.QtCore import QThread, pyqtSignal

logger = logging.getLogger(__name__)


@dataclass
class OcrResult:
    text: str
    bbox: Optional[list[list[float]]] = None
    original_text: Optional[str] = None
    processed_text: Optional[str] = None

    def __post_init__(self) -> None:
        # Zachovej původní text pro pozdější diakritizaci / exporty
        if self.original_text is None:
            self.original_text = self.text
        if self.processed_text is None:
            self.processed_text = self.text

    @property
    def display_text(self) -> str:
        """Text pro zobrazení/export – processed pokud existuje."""
        if self.processed_text is not None:
            return self.processed_text
        return self.text


LANG_MAP_EASYOCR = {
    "ces": "cs", "eng": "en", "deu": "de",
    "fra": "fr", "ita": "it", "pol": "pl",
}


def _extract_lang_code(lang_combo_text: str) -> str:
    return lang_combo_text.split()[0]


# ---------------------------------------------------------------------------
# EasyOCR – správa Readerů, cache a synchronizace
# ---------------------------------------------------------------------------

_EASYOCR_MODEL_DIR = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    ".easyocr_model",
)

# Centrální cache Readerů: easy_lang ("cs","en"…) -> Reader
_easyocr_readers: dict[str, object] = {}
# Stav per jazyk: "idle" | "preparing" | "ready" | "failed"
_easyocr_status: dict[str, str] = {}
_easyocr_errors: dict[str, str] = {}

# Sdílený lock + condition pro synchronizaci preloadu a _ensure_reader
_easyocr_lock = threading.Lock()
_easyocr_condition = threading.Condition(_easyocr_lock)


def _get_easyocr_model_dir() -> str:
    return _EASYOCR_MODEL_DIR


def _is_easyocr_model_cached(easy_lang: str) -> bool:
    """Ověří zda model pro daný jazyk již existuje v cache.
    EasyOCR stahuje do model_storage_directory; kontrola existence souborů
    slouží pouze pro log/diagnostiku a pro hlasové 'Stahuji model'.
    """
    try:
        if not os.path.isdir(_EASYOCR_MODEL_DIR):
            return False
        # craft je společný pro všechny jazyky
        craft_path = os.path.join(_EASYOCR_MODEL_DIR, "craft_mlt_25k.pth")
        has_craft = os.path.exists(craft_path)
        # jazykový model – EasyOCR pojmenovává např. cs_g2.pth, en_g2.pth
        lang_pattern = os.path.join(_EASYOCR_MODEL_DIR, f"{easy_lang}*.pth")
        has_lang = bool(glob.glob(lang_pattern))
        # také zkusit bez wildcard přímo (např. cs.pth)
        if not has_lang:
            has_lang = os.path.exists(os.path.join(_EASYOCR_MODEL_DIR, f"{easy_lang}.pth"))
        # pro češtinu někdy "cs_g2.pth", pro angličtinu "en_g2.pth"
        # Pokud chybí obojí ale adresář obsahuje cokoliv pro daný jazyk, považuj za cached
        if not has_lang:
            files = os.listdir(_EASYOCR_MODEL_DIR)
            has_lang = any(f.startswith(easy_lang) and f.endswith(".pth") for f in files)
        # minimálně craft + lang musí existovat, ale pokud chybí jen craft a lang existuje,
        # stále může být potřeba stáhnout craft – považuj za ne-cached
        return has_craft and has_lang
    except Exception:
        return False


def get_easyocr_status(easy_lang: str) -> str:
    with _easyocr_lock:
        return _easyocr_status.get(easy_lang, "idle")


def is_easyocr_ready(easy_lang: str) -> bool:
    with _easyocr_lock:
        return _easyocr_status.get(easy_lang) == "ready" and easy_lang in _easyocr_readers


def is_easyocr_preparing(easy_lang: str) -> bool:
    with _easyocr_lock:
        return _easyocr_status.get(easy_lang) == "preparing"


def is_easyocr_failed(easy_lang: str) -> bool:
    with _easyocr_lock:
        return _easyocr_status.get(easy_lang) == "failed"


def get_easyocr_reader(easy_lang: str):
    """Vrátí již existující Reader z paměti, nebo None.
    Thread-safe. Neprovádí inicializaci ani stahování.
    Konceptuálně: reader = get_easyocr_reader(language)
    """
    with _easyocr_lock:
        return _easyocr_readers.get(easy_lang)


def get_easyocr_error(easy_lang: str) -> str:
    with _easyocr_lock:
        return _easyocr_errors.get(easy_lang, "")


def _set_easyocr_status(easy_lang: str, status: str, error: str = "") -> None:
    with _easyocr_lock:
        _easyocr_status[easy_lang] = status
        if error:
            _easyocr_errors[easy_lang] = error
        elif status != "failed":
            _easyocr_errors.pop(easy_lang, None)
        # probuď čekající vlákna
        _easyocr_condition.notify_all()


def ensure_easyocr_reader(easy_lang: str):
    """Zajistí Reader pro daný jazyk. Blokující, volat pouze z worker vlákna.

    - vrátí již existující Reader z cache,
    - nebo vytvoří nový Reader (EasyOCR automaticky stáhne chybějící modely),
    - uloží do paměti a označí status ready,
    - synchronizováno _easyocr_lock/_condition aby nevznikly 2 paralelní inicializace.
    """
    # Rychlá cesta bez čekání
    with _easyocr_lock:
        if easy_lang in _easyocr_readers and _easyocr_status.get(easy_lang) == "ready":
            logger.info("EasyOCR reuse cached Reader lang=%s", easy_lang)
            return _easyocr_readers[easy_lang]
        # Pokud jiný thread právě připravuje stejný jazyk, počkej
        while _easyocr_status.get(easy_lang) == "preparing":
            logger.info("EasyOCR wait for preparing lang=%s", easy_lang)
            _easyocr_condition.wait(timeout=60)
            if easy_lang in _easyocr_readers and _easyocr_status.get(easy_lang) == "ready":
                return _easyocr_readers[easy_lang]
            # pokud se mezitím změnil stav na failed, vyskoč a zkus znovu vytvořit
            if _easyocr_status.get(easy_lang) == "failed":
                break
        if easy_lang in _easyocr_readers and _easyocr_status.get(easy_lang) == "ready":
            return _easyocr_readers[easy_lang]
        # Označ jako preparing
        _easyocr_status[easy_lang] = "preparing"
        _easyocr_condition.notify_all()

    cached = _is_easyocr_model_cached(easy_lang)
    logger.info("EasyOCR ensure_reader start lang=%s cached=%s dir=%s", easy_lang, cached, _EASYOCR_MODEL_DIR)
    if cached:
        logger.info("EasyOCR model already in cache lang=%s", easy_lang)
    else:
        logger.info("EasyOCR model will be downloaded lang=%s", easy_lang)

    try:
        reader = easyocr.Reader(
            [easy_lang],
            gpu=False,
            model_storage_directory=_EASYOCR_MODEL_DIR,
        )
        with _easyocr_lock:
            _easyocr_readers[easy_lang] = reader
            _easyocr_status[easy_lang] = "ready"
            _easyocr_errors.pop(easy_lang, None)
            _easyocr_condition.notify_all()
        logger.info("EasyOCR Reader ready lang=%s", easy_lang)
        # Synchronizuj i legacy singleton pro kompatibilitu
        try:
            EasyOCRThread._reader = reader
            EasyOCRThread._reader_lang = [easy_lang]
        except Exception:
            pass
        return reader
    except Exception as e:
        logger.exception("EasyOCR Reader failed lang=%s err=%s", easy_lang, e)
        with _easyocr_lock:
            _easyocr_status[easy_lang] = "failed"
            _easyocr_errors[easy_lang] = str(e)
            _easyocr_condition.notify_all()
        raise


def clear_easyocr_cache() -> None:
    """Pro testy – vymaže paměťovou cache (ne soubory na disku)."""
    with _easyocr_lock:
        _easyocr_readers.clear()
        _easyocr_status.clear()
        _easyocr_errors.clear()
        _easyocr_condition.notify_all()
    try:
        EasyOCRThread._reader = None
        EasyOCRThread._reader_lang = None
    except Exception:
        pass


class TesseractThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)
    ocr_results = pyqtSignal(list)
    diacritics_failed = pyqtSignal(str)

    def __init__(self, images: list[Image.Image], lang_code: str, diacritics_enabled: bool = False, raw_lang_code: str = "") -> None:
        super().__init__()
        self.images = images
        self.lang_code = _extract_lang_code(lang_code)
        self.raw_lang_code = raw_lang_code or lang_code
        self.diacritics_enabled = diacritics_enabled

    def run(self) -> None:
        full_text = ""
        results_list: list[list[OcrResult]] = []
        total = len(self.images)

        for i, img in enumerate(self.images):
            data = pytesseract.image_to_data(
                img, lang=self.lang_code, output_type=pytesseract.Output.DICT
            )

            page_results: list[OcrResult] = []
            n_boxes = len(data["text"])
            for j in range(n_boxes):
                conf = int(data["conf"][j])
                if conf > 0 and data["text"][j].strip():
                    x, y, w, h = (
                        data["left"][j], data["top"][j],
                        data["width"][j], data["height"][j],
                    )
                    bbox = [
                        [float(x), float(y)],
                        [float(x + w), float(y)],
                        [float(x + w), float(y + h)],
                        [float(x), float(y + h)],
                    ]
                    page_results.append(OcrResult(text=data["text"][j], bbox=bbox))

            page_text = pytesseract.image_to_string(img, lang=self.lang_code)

            # --- Volitelná diakritizace (mimo GUI vlákno, uvnitř workeru) ---
            try:
                from diacritics import should_diacritize, get_diacritizer
                if should_diacritize(self.raw_lang_code, self.diacritics_enabled):
                    diac = get_diacritizer()
                    corrected_page_text = diac.diacritize(page_text)
                    # per-word bbox zachován, text opraven
                    new_page_results: list[OcrResult] = []
                    for r in page_results:
                        orig = r.text
                        corr = diac.diacritize(orig)
                        new_page_results.append(OcrResult(text=corr, bbox=r.bbox, original_text=orig, processed_text=corr))
                    page_results = new_page_results
                    page_text = corrected_page_text
                else:
                    # vyplň original/processed identicky
                    page_results = [OcrResult(text=r.text, bbox=r.bbox, original_text=r.text, processed_text=r.text) for r in page_results]
            except Exception as e:
                try:
                    self.diacritics_failed.emit(str(e))
                except Exception:
                    pass
                page_results = [OcrResult(text=r.text, bbox=r.bbox, original_text=r.text, processed_text=r.text) for r in page_results]

            full_text += f"--- Stránka {i + 1} ---\n{page_text}\n\n"
            results_list.append(page_results)

            self.progress.emit(int((i + 1) / total * 100))

        self.ocr_results.emit(results_list)
        self.finished.emit(full_text)


class EasyOCRThread(QThread):
    _lock = _easyocr_lock
    _reader = None
    _reader_lang = None

    progress = pyqtSignal(int)
    finished = pyqtSignal(str)
    ocr_results = pyqtSignal(list)
    model_loading = pyqtSignal(int)
    diacritics_failed = pyqtSignal(str)

    def __init__(self, images: list[Image.Image], lang_code: str, diacritics_enabled: bool = False, raw_lang_code: str = "") -> None:
        super().__init__()
        self.images = images
        pyt_code = _extract_lang_code(lang_code)
        self.lang_code = LANG_MAP_EASYOCR.get(pyt_code, pyt_code[:2])
        self.raw_lang_code = raw_lang_code or lang_code
        self.diacritics_enabled = diacritics_enabled

    # dodatečný signál pro chybu inicializace (nepřipraveno)
    failed = pyqtSignal(str)

    def run(self) -> None:
        try:
            self._ensure_reader()
        except Exception as e:
            logger.exception("EasyOCRThread run failed in _ensure_reader lang=%s", self.lang_code)
            try:
                self.failed.emit(str(e))
            except Exception:
                pass
            # Ukončit s prázdným výsledkem aby QEventLoop nezůstal viset
            try:
                self.ocr_results.emit([])
            except Exception:
                pass
            try:
                self.finished.emit("")
            except Exception:
                pass
            return

        full_text = ""
        results_list: list[list[OcrResult]] = []
        total = len(self.images)

        for i, img in enumerate(self.images):
            img_np = np.array(img)
            raw_results = EasyOCRThread._reader.readtext(img_np, paragraph=True)

            page_text = ""
            page_results: list[OcrResult] = []
            for result in raw_results:
                if len(result) == 3:
                    bbox, text, _ = result
                    page_results.append(OcrResult(text=str(text), bbox=bbox))
                else:
                    word_results, _ = result
                    if not isinstance(word_results, list) or not word_results:
                        continue
                    texts = []
                    all_bboxes = []
                    for word in word_results:
                        if isinstance(word, (list, tuple)) and len(word) >= 3:
                            texts.append(str(word[1]))
                            all_bboxes.append(word[0])
                    text = " ".join(texts)
                    if not text.strip():
                        continue
                    if all_bboxes:
                        xs = [p[0] for b in all_bboxes for p in b]
                        ys = [p[1] for b in all_bboxes for p in b]
                        bbox = [
                            [min(xs), min(ys)], [max(xs), min(ys)],
                            [max(xs), max(ys)], [min(xs), max(ys)],
                        ]
                    else:
                        bbox = None
                    page_results.append(OcrResult(text=text, bbox=bbox))

                page_text += result[1] if len(result) == 3 else text + "\n"

            # --- Volitelná diakritizace (mimo GUI vlákno) ---
            try:
                from diacritics import should_diacritize, get_diacritizer
                if should_diacritize(self.raw_lang_code, self.diacritics_enabled):
                    diac = get_diacritizer()
                    corrected_page_text = diac.diacritize(page_text)
                    new_page_results: list[OcrResult] = []
                    for r in page_results:
                        orig = r.text
                        corr = diac.diacritize(orig)
                        new_page_results.append(OcrResult(text=corr, bbox=r.bbox, original_text=orig, processed_text=corr))
                    page_results = new_page_results
                    page_text = corrected_page_text
                else:
                    page_results = [OcrResult(text=r.text, bbox=r.bbox, original_text=r.text, processed_text=r.text) for r in page_results]
            except Exception as e:
                try:
                    self.diacritics_failed.emit(str(e))
                except Exception:
                    pass
                page_results = [OcrResult(text=r.text, bbox=r.bbox, original_text=r.text, processed_text=r.text) for r in page_results]

            full_text += f"--- Stránka {i + 1} ---\n{page_text}\n\n"
            results_list.append(page_results)

            self.progress.emit(int((i + 1) / total * 100))

        self.ocr_results.emit(results_list)
        self.finished.emit(full_text)

    def _ensure_reader(self) -> None:
        # Reuse centralizovaného správce – nevytváří duplicitní Reader
        # 1) rychlá cesta: Reader již v paměti
        existing = get_easyocr_reader(self.lang_code)
        if existing is not None and EasyOCRThread._reader is not None and EasyOCRThread._reader_lang == [self.lang_code]:
            # již nastaveno lokálně – rychlý reuse
            return
        if existing is not None:
            # reuse z centrální cache
            EasyOCRThread._reader = existing
            EasyOCRThread._reader_lang = [self.lang_code]
            logger.info("EasyOCRThread reuse central cache lang=%s", self.lang_code)
            return
        # 2) potřeba vytvořit – použij synchronizovaný ensure (blokuje jen tento worker)
        # Rozliš cached vs download pro přístupné hlášky
        cached_before = _is_easyocr_model_cached(self.lang_code)
        self.model_loading.emit(0)
        if not cached_before:
            logger.info("EasyOCRThread downloading model lang=%s", self.lang_code)
        else:
            logger.info("EasyOCRThread loading cached model lang=%s", self.lang_code)
        try:
            reader = ensure_easyocr_reader(self.lang_code)
            EasyOCRThread._reader = reader
            EasyOCRThread._reader_lang = [self.lang_code]
            self.model_loading.emit(100)
        except Exception as e:
            # Propaguj výjimku – run() ji nechá spadnout do error handlingu volajícího
            # ale zároveň emituj 100 aby progress dialog nezůstal viset
            logger.exception("EasyOCRThread _ensure_reader failed lang=%s", self.lang_code)
            self.model_loading.emit(100)
            raise


class EasyOCRPreloadThread(QThread):
    """Worker pro přípravu EasyOCR na pozadí bez blokování GUI.

    Používá stejný Lock/Condition jako EasyOCRThread._ensure_reader,
    takže nevznikne duplicitní stahování/inicializace.
    """

    started = pyqtSignal(str)        # lang easy (cs)
    downloading = pyqtSignal(str)    # lang – model chybí v cache, bude stahování
    finished_ok = pyqtSignal(str)    # lang
    failed = pyqtSignal(str, str)    # lang, error_msg
    progress_msg = pyqtSignal(str)   # přístupná textová hláška

    def __init__(self, lang_code_easy: str) -> None:
        super().__init__()
        self.lang_code = lang_code_easy

    def run(self) -> None:
        try:
            self.started.emit(self.lang_code)
            self.progress_msg.emit("Připravuji OCR.")
            logger.info("EasyOCRPreloadThread start lang=%s", self.lang_code)
            cached = _is_easyocr_model_cached(self.lang_code)
            if not cached:
                self.downloading.emit(self.lang_code)
                self.progress_msg.emit("Stahuji model pro EasyOCR.")
                logger.info("EasyOCRPreloadThread downloading lang=%s", self.lang_code)
            else:
                logger.info("EasyOCRPreloadThread loading cached lang=%s", self.lang_code)
            ensure_easyocr_reader(self.lang_code)
            self.finished_ok.emit(self.lang_code)
            self.progress_msg.emit("EasyOCR je připraven.")
            logger.info("EasyOCRPreloadThread finished lang=%s", self.lang_code)
        except Exception as e:
            msg = str(e)
            logger.exception("EasyOCRPreloadThread failed lang=%s err=%s", self.lang_code, e)
            self.failed.emit(self.lang_code, msg)
            self.progress_msg.emit("EasyOCR se nepodařilo připravit. Zkontrolujte připojení k internetu a zkuste to znovu.")


def create_ocr_thread(
    engine: str, images: list[Image.Image], lang_code: str, diacritics_enabled: bool = False
) -> TesseractThread | EasyOCRThread:
    if engine == "EasyOCR":
        return EasyOCRThread(images, lang_code, diacritics_enabled=diacritics_enabled, raw_lang_code=lang_code)
    return TesseractThread(images, lang_code, diacritics_enabled=diacritics_enabled, raw_lang_code=lang_code)
