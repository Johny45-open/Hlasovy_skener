"""Import existujicich PDF souboru se selektivnim OCR (MVP).

Tok:  PDF import -> analyza textove vrstvy -> stranky potrebujici OCR
       -> render potrebnych stran -> OCR -> overlay textove vrstvy
       do KOPIE originalu -> novy <nazev>_OCR.pdf.

Dulezita rozhodnuti (viz audit pred BUILD rezimem):

1. ZADNY QThread uvnitr QThread.
   Existujici ``TesseractThread`` / ``EasyOCRThread`` (ocr_engine.py) jsou
   samy o sobe QThread a nelze je bezpecne spoustet z ``run()`` jineho
   QThread (thread-affinity, event-loop). Proto tento modul obsahuje
   synchronni OCR funkce ``tesseract_ocr_images_sync()`` /
   ``easyocr_ocr_images_sync()``, ktere jsou zamerne zrcadlem logiky
   ``TesseractThread.run()`` / ``EasyOCRThread.run()`` (stejne parametry
   pytesseract / easyocr, stejny parsing bbox, stejna diakritizace,
   stejna cache ``ensure_easyocr_reader``). Sdili typy (``OcrResult``),
   jazykovou mapu i fontovou vrstvu - nevznika druhy OCR system,
   jen synchronni adapter pro pouziti ve workeru.

2. Textova vrstva se NEREIMPLEMENTUJE. ``overlay_ocr_layer()`` pouziva
   existujici ``Skener._insert_ocr_textbox_invisible()``,
   ``_resolve_ocr_pdf_fontfile()`` a ``_ocr_pdf_initial_fontsize()``
   (lazy import uvnitr funkce, aby nevznikl cyklicky import
   Skener <-> pdf_import). Stejny font (DejaVuSans), stejny
   render_mode=3, stejny shrink-loop i fallback.

3. Mixed PDF se NIKDY nerasterizuje cele. Overlay zapisuje pouze do
   stran urcenych k OCR; puvodni textove strany zustavaji nedotcene.

4. Worker (``PdfImportWorker``) neprovadi ZADNE GUI operace. Komunikuje
   s GUI vlaknem vyhradne pres Qt signaly. Dialogy, speak() a
   ``_confirm_overwrite()`` resi vzdy GUI vlakno (Skener.py).
"""

from __future__ import annotations

import logging
import os
from dataclasses import dataclass, field
from typing import Callable, Optional

import fitz
import numpy as np
import pytesseract
from PIL import Image
from PyQt6.QtCore import QThread, pyqtSignal

from ocr_engine import OcrResult, LANG_MAP_EASYOCR, ensure_easyocr_reader

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Pojmenovane konstanty (laditelne bez zasahu do logiky)
# ---------------------------------------------------------------------------

#: Rozliseni renderu PDF stran pro OCR. Shoda s defaultem dpi_combo (300).
PDF_IMPORT_RENDER_DPI = 300

#: Strana je povazovana za prohledavatelnou, pokud ma alespon tolik znaku...
PDF_SEARCHABLE_MIN_CHARS = 100

#: ...nebo alespon tolik slov SOUCASNE s minimalnim poctem alfanumerickych znaku.
PDF_SEARCHABLE_MIN_WORDS = 20
PDF_SEARCHABLE_MIN_ALNUM = 50

#: Pripona vystupniho souboru (``faktura.pdf`` -> ``faktura_OCR.pdf``).
PDF_IMPORT_OUTPUT_SUFFIX = "_OCR"

#: Rotace stran, pro ktere umi overlay spocitat souradnice (0 = primy
#: prepocet, 180 = zrcadleny prepocet). 90/270 se bezpecne odmitnou.
PDF_IMPORT_SUPPORTED_ROTATIONS = (0, 180)


# ---------------------------------------------------------------------------
# Vyjimky (per-file granularita - chyba jednoho PDF neukonci davku)
# ---------------------------------------------------------------------------

class PdfImportError(Exception):
    """Zakladni chyba importu jednoho PDF."""


class CorruptPdfError(PdfImportError):
    """PDF nelze otevrit / je poskozene."""


class EncryptedPdfError(PdfImportError):
    """PDF je chranene heslem (bez poskytnuteho hesla nezpracovatelne)."""


class UnsupportedRotationError(PdfImportError):
    """Strana urcena k OCR ma nepodporovanou rotaci (90/270)."""


class PageRenderError(PdfImportError):
    """Stranku se nepodarilo vyrenderovat do obrazku."""


class OcrEngineError(PdfImportError):
    """OCR selhalo (chybejici engine / jazyk, chyba modelu)."""


class FontError(PdfImportError):
    """Chybi Unicode font pro textovou vrstvu."""


class SaveError(PdfImportError):
    """Vystup se nepodarilo zapsat (prava, cesta, ...)."""


class ValidationError(PdfImportError):
    """Vysledne PDF neproslo validaci (povazovano za chybu exportu)."""


class CancelledError(PdfImportError):
    """Zpracovani zruseno uzivatelem (vnitrni ridici vyjimka)."""


# ---------------------------------------------------------------------------
# Datove struktury
# ---------------------------------------------------------------------------

@dataclass
class PdfPageNeed:
    """Vysledek analyzy jedne strany (cislovani od 1)."""
    page_no: int
    needs_ocr: bool
    char_count: int = 0
    word_count: int = 0
    alnum_count: int = 0
    block_count: int = 0
    image_count: int = 0
    rotation: int = 0


@dataclass
class PdfAnalysis:
    """Vysledek analyzy jednoho PDF souboru."""
    path: str
    page_count: int = 0
    pages: list[PdfPageNeed] = field(default_factory=list)
    error: str = ""  # neprazdne = soubor nelze zpracovat (duvod)

    @property
    def ok(self) -> bool:
        return not self.error

    @property
    def pages_needing_ocr(self) -> list[int]:
        return [p.page_no for p in self.pages if p.needs_ocr]


@dataclass
class ImportJob:
    """Prace pro worker: jeden zdrojovy soubor -> jeden vystup."""
    src_path: str
    out_path: str
    pages_to_ocr: list[int] = field(default_factory=list)  # 1-based
    engine: str = "Tesseract"
    lang_raw: str = "ces (Čeština)"
    diacritics_enabled: bool = False


@dataclass
class JobResult:
    """Vysledek jednoho jobu (status: ok | skipped | failed | cancelled)."""
    src_path: str
    out_path: str = ""
    status: str = "failed"
    message: str = ""
    ocr_pages: list[int] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Analyza PDF (rychla, ciste textova extrakce - vhodna i pro GUI vlakno)
# ---------------------------------------------------------------------------

def _classify_page(page: "fitz.Page", page_no: int) -> PdfPageNeed:
    """Klasifikuje jednu stranu podle konzervativni heuristiky.

    Zamerne NENI pouzito samotne ``page.get_text() == ""``. Strana je
    povazovana za prohledavatelnou, pokud ma dostatek souvisleho textu;
    kratke fragmenty (razitko, zahlavi skenu) bezpecne padaji do NEEDS_OCR
    (radsi OCR navic nez ztraceny obsah - overlay textovou stranu neznici).
    """
    try:
        text = page.get_text("text") or ""
    except Exception:
        text = ""
    stripped = text.strip()
    words = stripped.split()
    alnum = sum(1 for c in stripped if c.isalnum())
    try:
        blocks = page.get_text("blocks") or []
    except Exception:
        blocks = []
    try:
        images = page.get_images(full=True) or []
    except Exception:
        images = []
    try:
        rotation = int(page.rotation or 0) % 360
    except Exception:
        rotation = 0

    searchable = (
        len(stripped) >= PDF_SEARCHABLE_MIN_CHARS
        or (len(words) >= PDF_SEARCHABLE_MIN_WORDS
            and alnum >= PDF_SEARCHABLE_MIN_ALNUM)
    )
    return PdfPageNeed(
        page_no=page_no,
        needs_ocr=not searchable,
        char_count=len(stripped),
        word_count=len(words),
        alnum_count=alnum,
        block_count=len(blocks),
        image_count=len(images),
        rotation=rotation,
    )


def analyze_pdf(path: str) -> PdfAnalysis:
    """Analyzuje PDF a vrati pozadavek OCR pro kazdou stranu.

    Nikdy nevyhodi - neotevritelne / zaheslovane PDF vrati
    ``PdfAnalysis(error=...)`` se srozumitelnym duvodem.
    """
    if not os.path.isfile(path):
        return PdfAnalysis(path=path, error=f"Soubor neexistuje: {path}")
    doc = None
    try:
        try:
            doc = fitz.open(path)
        except Exception as e:
            return PdfAnalysis(path=path, error=f"PDF nelze otevřít (možná poškozený soubor): {e}")
        try:
            if getattr(doc, "needs_pass", False) or getattr(doc, "is_encrypted", False):
                # Pokus o prazdne heslo - nektera PDF ho akceptuji.
                try:
                    authenticated = doc.authenticate("")
                except Exception:
                    authenticated = False
                if not authenticated and getattr(doc, "needs_pass", False):
                    return PdfAnalysis(
                        path=path,
                        error="PDF je chráněno heslem. Zadejte heslo v prohlížeči PDF, uložte nechráněnou kopii a importujte ji.",
                    )
        except EncryptedPdfError:
            raise
        except Exception:
            pass
        try:
            page_count = len(doc)
        except Exception as e:
            return PdfAnalysis(path=path, error=f"PDF nemá čitelné stránky: {e}")
        if page_count == 0:
            return PdfAnalysis(path=path, page_count=0, error="PDF neobsahuje žádné stránky.")
        pages = [_classify_page(doc[i], i + 1) for i in range(page_count)]
        return PdfAnalysis(path=path, page_count=page_count, pages=pages)
    finally:
        if doc is not None:
            try:
                doc.close()
            except Exception:
                pass


def _stran_word(n: int) -> str:
    """Spravne sklonovani: 1 strana, 2-4 strany, 5+ stran."""
    if n == 1:
        return "strana"
    if 2 <= n <= 4:
        return "strany"
    return "stran"


def format_analysis_summary(analysis: PdfAnalysis) -> str:
    """Jednoradkovy (prip. dvouvetny) souhrn pro pristupny dialog."""
    base = os.path.basename(analysis.path)
    if analysis.error:
        return f"Dokument {base} nelze zpracovat: {analysis.error}"
    need = analysis.pages_needing_ocr
    n = analysis.page_count
    if not need:
        amount = "1 stranu" if n == 1 else f"{n} {_stran_word(n)}"
        return (f"Dokument {base} má {amount}. "
                f"Všechny strany již obsahují použitelný text. OCR není potřeba.")
    lst = ", ".join(str(p) for p in need)
    return (f"Dokument {base} má {n} {_stran_word(n)}. "
            f"OCR je potřeba na {len(need)} stranách: {lst}.")


# ---------------------------------------------------------------------------
# Cesty a davkove planovani (ciste funkce - testovatelne bez GUI)
# ---------------------------------------------------------------------------

def default_output_path(src_path: str) -> str:
    """Vychozi vystup ``<nazev>_OCR.pdf`` ve stejne slozce jako zdroj.

    Nikdy nevrati cestu shodnou se zdrojem (i ``*_OCR.pdf`` dostane
    dalsi suffix). Zdroj se tak nemuze prepsat ani omylem.
    """
    directory = os.path.dirname(os.path.abspath(src_path))
    stem, _ext = os.path.splitext(os.path.basename(src_path))
    candidate = os.path.join(directory, f"{stem}{PDF_IMPORT_OUTPUT_SUFFIX}.pdf")
    if os.path.abspath(candidate) == os.path.abspath(src_path):
        candidate = os.path.join(directory, f"{stem}{PDF_IMPORT_OUTPUT_SUFFIX}_2.pdf")
    return candidate


def build_jobs(
    analyses: list[PdfAnalysis],
    engine: str,
    lang_raw: str,
    diacritics_enabled: bool,
    overwrite_allowed: dict[str, bool] | None = None,
) -> tuple[list[ImportJob], list[JobResult]]:
    """Sestavi joby pro soubory potrebujici OCR.

    ``overwrite_allowed`` mapuje out_path -> True/False (rozhodnuti
    z GUI ``_confirm_overwrite``). Soubor s existujicim vystupem bez
    souhlasu je ``skipped`` (davka pokracuje dalsimi soubory).
    Vraci (jobs, skipped_results).
    """
    overwrite_allowed = overwrite_allowed or {}
    jobs: list[ImportJob] = []
    skipped: list[JobResult] = []
    for analysis in analyses:
        if not analysis.ok:
            skipped.append(JobResult(
                src_path=analysis.path, status="skipped",
                message=f"Přeskočeno - nelze analyzovat: {analysis.error}"))
            continue
        need = analysis.pages_needing_ocr
        if not need:
            continue  # OCR neni potreba, nic se neuklada
        out = default_output_path(analysis.path)
        if os.path.exists(out) and not overwrite_allowed.get(out, False):
            skipped.append(JobResult(
                src_path=analysis.path, out_path=out, status="skipped",
                message="Přeskočeno - výstup již existuje a přepsání nebylo potvrzeno."))
            continue
        jobs.append(ImportJob(
            src_path=analysis.path, out_path=out, pages_to_ocr=need,
            engine=engine, lang_raw=lang_raw,
            diacritics_enabled=diacritics_enabled))
    return jobs, skipped


# ---------------------------------------------------------------------------
# Render PDF strany -> PIL (DPI 300, RGB)
# ---------------------------------------------------------------------------

def render_page_to_image(page: "fitz.Page", dpi: int = PDF_IMPORT_RENDER_DPI) -> Image.Image:
    """Vyrenderuje stranu do RGB PIL obrazku pro OCR.

    Stejne DPI jako doporuceny sken (300). Prevod bbox zpet do PDF bodu
    je pak konzistentni se soucasnym exportem (``pixel / zoom``,
    kde ``zoom = dpi / 72`` == inverze ``x / dpi * 72``).
    """
    if dpi is None or int(dpi) <= 0:
        dpi = PDF_IMPORT_RENDER_DPI
    zoom = float(dpi) / 72.0
    try:
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    except Exception as e:
        raise PageRenderError(f"Stránku se nepodařilo vyrenderovat: {e}")
    try:
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    except Exception as e:
        raise PageRenderError(f"Renderovanou stránku nelze převést na obrázek: {e}")
    return img.copy()


def pixel_bbox_to_pdf_rect(
    bbox,
    zoom: float,
    rotation: int,
    page_width_pt: float,
    page_height_pt: float,
) -> Optional["fitz.Rect"]:
    """Prevede OCR bbox (pixely renderu) na PDF rect (body).

    Podporuje rotace 0 a 180 (overena transformace). Pro 90/270 vyhodi
    ``UnsupportedRotationError`` - souradnice se NEHADAJI.
    Vraci None pro degenerovany bbox (volajici pouzije nobbox fallback).
    """
    rot = int(rotation or 0) % 360
    if rot not in PDF_IMPORT_SUPPORTED_ROTATIONS:
        raise UnsupportedRotationError(
            f"Strana má rotaci {rot}°. Podporovány jsou pouze rotace 0° a 180° - "
            f"soubor nebude falešně OCRován.")
    try:
        xs = [float(p[0]) for p in bbox]
        ys = [float(p[1]) for p in bbox]
    except Exception:
        return None
    if not xs or not ys:
        return None
    xa, xb = min(xs) / zoom, max(xs) / zoom
    ya, yb = min(ys) / zoom, max(ys) / zoom
    if xb <= xa or yb <= ya:
        return None
    if rot == 180:
        xa, xb = page_width_pt - xb, page_width_pt - xa
        ya, yb = page_height_pt - yb, page_height_pt - ya
    return fitz.Rect(xa, ya, xb, yb)


# ---------------------------------------------------------------------------
# Synchronni OCR adapter (zrcadlo TesseractThread / EasyOCRThread BEZ QThread)
# ---------------------------------------------------------------------------

def _apply_diacritics_sync(
    page_results: list[OcrResult], page_text: str,
    lang_raw: str, diacritics_enabled: bool,
) -> tuple[list[OcrResult], str]:
    """Diakritizace shodna s OCR thready (pouze ces + zapnuta volba)."""
    from diacritics import should_diacritize, get_diacritizer
    if not should_diacritize(lang_raw, diacritics_enabled):
        return ([OcrResult(text=r.text, bbox=r.bbox,
                           original_text=r.text, processed_text=r.text)
                 for r in page_results], page_text)
    try:
        diac = get_diacritizer()
        new_results = []
        for r in page_results:
            corr = diac.diacritize(r.text)
            new_results.append(OcrResult(text=corr, bbox=r.bbox,
                                        original_text=r.text, processed_text=corr))
        return new_results, diac.diacritize(page_text)
    except Exception as e:
        logger.warning("Oprava diakritiky selhala, zachovan puvodni text: %s", e)
        return ([OcrResult(text=r.text, bbox=r.bbox,
                           original_text=r.text, processed_text=r.text)
                 for r in page_results], page_text)


def tesseract_ocr_images_sync(
    images: list[Image.Image],
    lang_raw: str,
    diacritics_enabled: bool = False,
    progress_cb: Optional[Callable[[int, int], None]] = None,
    is_cancelled: Optional[Callable[[], bool]] = None,
) -> list[list[OcrResult]]:
    """Synchronni ekvivalent ``TesseractThread.run()`` pro worker importu."""
    lang_code = (lang_raw or "").split()[0].strip() or "ces"
    all_pages: list[list[OcrResult]] = []
    total = len(images)
    try:
        import pytesseract as _pt
    except Exception as e:
        raise OcrEngineError(f"Tesseract není dostupný: {e}")
    for i, img in enumerate(images):
        if is_cancelled is not None and is_cancelled():
            raise CancelledError("OCR zrušeno uživatelem.")
        try:
            data = _pt.image_to_data(img, lang=lang_code,
                                     output_type=_pt.Output.DICT)
        except Exception as e:
            msg = str(e)
            if "Failed loading language" in msg or "No such file" in msg:
                raise OcrEngineError(
                    f"Tesseract nemá nainstalovaný jazyk '{lang_code}'. "
                    f"Nainstalujte jazyková data a opakujte. Detail: {msg}")
            if "tesseract is not installed" in msg.lower() or "No such file or directory" in msg:
                raise OcrEngineError(
                    f"Tesseract OCR engine není nainstalován nebo není v PATH. Detail: {msg}")
            raise OcrEngineError(f"OCR selhalo na straně {i + 1}: {msg}")
        page_results: list[OcrResult] = []
        try:
            n_boxes = len(data["text"])
            for j in range(n_boxes):
                try:
                    conf = int(float(data["conf"][j]))
                except Exception:
                    conf = -1
                word = data["text"][j]
                if conf > 0 and word and str(word).strip():
                    x, y, w, h = (data["left"][j], data["top"][j],
                                  data["width"][j], data["height"][j])
                    bbox = [[float(x), float(y)], [float(x + w), float(y)],
                            [float(x + w), float(y + h)], [float(x), float(y + h)]]
                    page_results.append(OcrResult(text=str(word), bbox=bbox))
            page_text = _pt.image_to_string(img, lang=lang_code)
        except OcrEngineError:
            raise
        except Exception as e:
            raise OcrEngineError(f"OCR selhalo na straně {i + 1}: {e}")
        page_results, _ = _apply_diacritics_sync(
            page_results, page_text, lang_raw, diacritics_enabled)
        all_pages.append(page_results)
        if progress_cb is not None:
            progress_cb(i + 1, total)
    return all_pages


def easyocr_ocr_images_sync(
    images: list[Image.Image],
    lang_raw: str,
    diacritics_enabled: bool = False,
    progress_cb: Optional[Callable[[int, int], None]] = None,
    is_cancelled: Optional[Callable[[], bool]] = None,
) -> list[list[OcrResult]]:
    """Synchronni ekvivalent ``EasyOCRThread.run()`` pro worker importu."""
    pyt_code = (lang_raw or "").split()[0].strip() or "ces"
    easy_lang = LANG_MAP_EASYOCR.get(pyt_code, pyt_code[:2])
    try:
        reader = ensure_easyocr_reader(easy_lang)
    except Exception as e:
        raise OcrEngineError(
            f"EasyOCR se nepodařilo připravit (jazyk '{easy_lang}'). "
            f"Zkontrolujte připojení k internetu a opakujte. Detail: {e}")
    all_pages: list[list[OcrResult]] = []
    total = len(images)
    for i, img in enumerate(images):
        if is_cancelled is not None and is_cancelled():
            raise CancelledError("OCR zrušeno uživatelem.")
        try:
            img_np = np.array(img)
            raw_results = reader.readtext(img_np, paragraph=True)
        except Exception as e:
            raise OcrEngineError(f"EasyOCR selhalo na straně {i + 1}: {e}")
        page_text = ""
        page_results: list[OcrResult] = []
        for result in raw_results:
            if len(result) == 3:
                bbox, text, _ = result
                page_results.append(OcrResult(text=str(text), bbox=bbox))
                page_text += result[1]
            else:
                word_results, _ = result
                if not isinstance(word_results, list) or not word_results:
                    continue
                texts, all_bboxes = [], []
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
                    bbox = [[min(xs), min(ys)], [max(xs), min(ys)],
                            [max(xs), max(ys)], [min(xs), max(ys)]]
                else:
                    bbox = None
                page_results.append(OcrResult(text=text, bbox=bbox))
                page_text += text + "\n"
        page_results, _ = _apply_diacritics_sync(
            page_results, page_text, lang_raw, diacritics_enabled)
        all_pages.append(page_results)
        if progress_cb is not None:
            progress_cb(i + 1, total)
    return all_pages


def ocr_images_sync(
    images: list[Image.Image],
    engine: str,
    lang_raw: str,
    diacritics_enabled: bool = False,
    progress_cb: Optional[Callable[[int, int], None]] = None,
    is_cancelled: Optional[Callable[[], bool]] = None,
) -> list[list[OcrResult]]:
    """Rozcestnik synchronniho OCR pro import (Tesseract / EasyOCR)."""
    if engine == "EasyOCR":
        return easyocr_ocr_images_sync(images, lang_raw, diacritics_enabled,
                                       progress_cb, is_cancelled)
    return tesseract_ocr_images_sync(images, lang_raw, diacritics_enabled,
                                     progress_cb, is_cancelled)


# ---------------------------------------------------------------------------
# Overlay OCR vrstvy do KOPIE originalu (textove strany nedotceny)
# ---------------------------------------------------------------------------

def overlay_ocr_layer(
    src_path: str,
    ocr_by_page: dict[int, list[OcrResult]],
    dpi: int = PDF_IMPORT_RENDER_DPI,
    out_path: str = "",
    progress_cb: Optional[Callable[[int, int], None]] = None,
    is_cancelled: Optional[Callable[[], bool]] = None,
) -> dict:
    """Prida neviditelnou OCR vrstvu pouze na zadane strany kopie PDF.

    Zdrojovy soubor nikdy neprepise (kontroluje ``out != src``).
    Pouziva existujici ``Skener._insert_ocr_textbox_invisible()`` -
    zadna druha implementace vrstvy. Strany mimo ``ocr_by_page``
    zustanou bitove nedotcene (pouze se prekopiruji).
    """
    # Lazy import - zamezi cyklickemu importu Skener <-> pdf_import.
    try:
        from Skener import (
            _insert_ocr_textbox_invisible,
            _resolve_ocr_pdf_fontfile,
            _ocr_pdf_initial_fontsize,
            _OCR_PDF_FONTNAME,
        )
    except Exception as e:
        raise FontError(f"Nelze načíst mechanismus textové vrstvy: {e}")

    if not out_path:
        raise SaveError("Není zadána výstupní cesta.")
    if os.path.abspath(out_path) == os.path.abspath(src_path):
        raise SaveError("Výstupní soubor nesmí být shodný se zdrojovým PDF.")
    if dpi is None or int(dpi) <= 0:
        dpi = PDF_IMPORT_RENDER_DPI
    dpi = int(dpi)
    zoom = float(dpi) / 72.0

    fontfile = _resolve_ocr_pdf_fontfile()
    if fontfile is None:
        raise FontError(
            "Pro PDF s rozpoznaným textem chybí Unicode font "
            "(assets/fonts/DejaVuSans.ttf). Přeinstalujte aplikaci.")

    stats = {"pages": 0, "inserted": 0, "shrunk": 0,
             "fallback": 0, "failed": 0, "nobbox": 0}
    try:
        doc = fitz.open(src_path)
    except Exception as e:
        raise CorruptPdfError(f"PDF nelze otevřít: {e}")
    try:
        if getattr(doc, "needs_pass", False):
            raise EncryptedPdfError("PDF je chráněno heslem.")
        targets = sorted(ocr_by_page.keys())
        total = len(targets)
        for done, page_no in enumerate(targets):
            if is_cancelled is not None and is_cancelled():
                raise CancelledError("Ukládání zrušeno uživatelem.")
            if page_no < 1 or page_no > len(doc):
                raise SaveError(f"Strana {page_no} v PDF neexistuje.")
            page = doc[page_no - 1]
            rotation = int(page.rotation or 0) % 360
            if rotation not in PDF_IMPORT_SUPPORTED_ROTATIONS:
                raise UnsupportedRotationError(
                    f"Strana {page_no} má rotaci {rotation}°. "
                    f"Podporovány jsou pouze rotace 0° a 180° - "
                    f"vrstva by byla umístěna falešně, soubor proto neukládám.")
            try:
                page.insert_font(fontname=_OCR_PDF_FONTNAME, fontfile=fontfile)
            except Exception as e:
                raise FontError(f"Registrace Unicode fontu selhala: {e}")
            stats["pages"] += 1
            nobbox_texts: list[str] = []
            for item in ocr_by_page.get(page_no, []):
                txt = getattr(item, "display_text", None)
                if txt is None:
                    txt = getattr(item, "text", "")
                if callable(txt):
                    try:
                        txt = txt()
                    except Exception:
                        txt = ""
                if not txt or not str(txt).strip():
                    continue
                txt = str(txt)
                bbox = getattr(item, "bbox", None)
                rect = None
                if bbox is not None:
                    try:
                        rect = pixel_bbox_to_pdf_rect(
                            bbox, zoom, rotation,
                            page.rect.width, page.rect.height)
                    except UnsupportedRotationError:
                        raise
                    except Exception as e:
                        logger.warning("Neplatny bbox na strane %d (%s) - pouzit fallback.", page_no, e)
                        rect = None
                if rect is None:
                    # Bezpecny fallback: text se NESMI tiho zahodit.
                    # Strategie je jednoznacna: spodni pruh strany (stejny
                    # vzor jako build_searchable_pdf), zalogovano.
                    nobbox_texts.append(txt)
                    continue
                res = _insert_ocr_textbox_invisible(
                    page, rect, txt, _OCR_PDF_FONTNAME,
                    _ocr_pdf_initial_fontsize(rect))
                if res in ("ok", "shrunk", "fallback"):
                    stats["inserted"] += 1
                if res == "shrunk":
                    stats["shrunk"] += 1
                elif res == "fallback":
                    stats["fallback"] += 1
                elif res == "failed":
                    stats["failed"] += 1
            if nobbox_texts:
                stats["nobbox"] += len(nobbox_texts)
                logger.warning("Strana %d ma %d polozek bez bbox - fallback do spodniho pruhu.",
                               page_no, len(nobbox_texts))
                strip = fitz.Rect(36, page.rect.height - 72,
                                  page.rect.width - 36, page.rect.height - 36)
                joined = "\n".join(nobbox_texts)
                res = _insert_ocr_textbox_invisible(
                    page, strip, joined, _OCR_PDF_FONTNAME, 8.0)
                if res in ("ok", "shrunk", "fallback"):
                    stats["inserted"] += len(nobbox_texts)
                if res == "shrunk":
                    stats["shrunk"] += 1
                elif res == "fallback":
                    stats["fallback"] += 1
                elif res == "failed":
                    stats["failed"] += len(nobbox_texts)
            if progress_cb is not None:
                progress_cb(done + 1, total)
        try:
            doc.save(out_path, garbage=4, deflate=True)
        except Exception as e:
            raise SaveError(f"Výstup se nepodařilo uložit ({out_path}): {e}")
    finally:
        try:
            doc.close()
        except Exception:
            pass
    if stats["failed"]:
        logger.error("Overlay: %d textu se nepodarilo vlozit: %s", stats["failed"], out_path)
    return stats


# ---------------------------------------------------------------------------
# Validace vysledneho PDF (selhani = chyba exportu, nikdy "dokonceno")
# ---------------------------------------------------------------------------

def validate_imported_pdf(
    out_path: str,
    expected_pages: int,
    ocr_by_page: dict[int, list[OcrResult]],
    original_texts: dict[int, str],
) -> list[str]:
    """Znovu otevre vystup a overi pocet stran, OCR text i puvodni texty.

    Vraci seznam chyb (prazdny = OK). Kontroluje i nahradni ``?`` misto
    ceskych znaku (regrese fontu helv).
    """
    errors: list[str] = []
    if not os.path.isfile(out_path):
        return [f"Výstupní soubor nevznikl: {out_path}"]
    doc = None
    try:
        try:
            doc = fitz.open(out_path)
        except Exception as e:
            return [f"Výstupní PDF nelze otevřít (validace selhala): {e}"]
        if len(doc) != expected_pages:
            errors.append(
                f"Výstup má {len(doc)} stran, očekáváno {expected_pages}.")
        for page_no in sorted(ocr_by_page.keys()):
            if page_no < 1 or page_no > len(doc):
                errors.append(f"OCR strana {page_no} ve výstupu chybí.")
                continue
            try:
                text = doc[page_no - 1].get_text("text") or ""
            except Exception as e:
                errors.append(f"Text strany {page_no} nelze přečíst: {e}")
                continue
            if not text.strip():
                errors.append(
                    f"OCR strana {page_no} je po uložení prázdná - vrstva se neuložila.")
                continue
            # Overeni klicovych slov z OCR (normalizace whitespace).
            flat = " ".join(text.split())
            for item in ocr_by_page[page_no]:
                src_txt = getattr(item, "display_text", None) or getattr(item, "text", "")
                if not src_txt or not str(src_txt).strip():
                    continue
                for word in str(src_txt).split():
                    w = word.strip()
                    if len(w) >= 4 and w.isalnum() and w not in flat:
                        errors.append(
                            f"Na OCR straně {page_no} chybí očekávané slovo {w!r} - "
                            f"vrstva je neúplná.")
                        break
                else:
                    continue
                break
            # Detekce nahradnich '?' misto diakritiky.
            src_all = " ".join(
                str(getattr(i, "display_text", None) or getattr(i, "text", ""))
                for i in ocr_by_page[page_no])
            if "?" in text and "?" not in src_all:
                diac = [c for c in "ěščřžýáíéúůďťňó" if c in src_all.lower()]
                if diac:
                    errors.append(
                        f"Na straně {page_no} jsou náhradní '?' místo českých znaků - "
                        f"poškozená textová vrstva (font).")
        for page_no, orig in original_texts.items():
            if page_no in ocr_by_page:
                continue
            if page_no < 1 or page_no > len(doc):
                continue
            try:
                now = doc[page_no - 1].get_text("text") or ""
            except Exception as e:
                errors.append(f"Původní stranu {page_no} nelze přečíst: {e}")
                continue
            if now != orig:
                errors.append(
                    f"Původní textová strana {page_no} byla změněna - "
                    f"musí zůstat nedotčená.")
    finally:
        if doc is not None:
            try:
                doc.close()
            except Exception:
                pass
    return errors


# ---------------------------------------------------------------------------
# Zpracovani jednoho jobu (volano z workeru, testovatelne primo)
# ---------------------------------------------------------------------------

def process_single_job(
    job: ImportJob,
    on_page: Optional[Callable[[str], None]] = None,
    on_pages_done: Optional[Callable[[int], None]] = None,
    is_cancelled: Optional[Callable[[], bool]] = None,
) -> JobResult:
    """Sekvencne zpracuje jeden soubor: render -> OCR -> overlay -> validace."""
    cancelled = is_cancelled() if is_cancelled is not None else False
    if cancelled:
        return JobResult(src_path=job.src_path, out_path=job.out_path,
                         status="cancelled", message="Zrušeno uživatelem.")
    if os.path.abspath(job.out_path) == os.path.abspath(job.src_path):
        return JobResult(src_path=job.src_path, out_path=job.out_path,
                         status="failed",
                         message="Odmítnuto: výstup by přepsal zdrojový soubor.")
    doc = None
    try:
        try:
            doc = fitz.open(job.src_path)
        except Exception as e:
            raise CorruptPdfError(f"PDF nelze otevřít (možná poškozený soubor): {e}")
        try:
            if getattr(doc, "needs_pass", False):
                raise EncryptedPdfError("PDF je chráněno heslem.")
            # Rychla predkontrola rotaci - selhat drive nez probehne drage OCR.
            for page_no in job.pages_to_ocr:
                if page_no < 1 or page_no > len(doc):
                    raise SaveError(f"Strana {page_no} v PDF neexistuje.")
                rot = int(doc[page_no - 1].rotation or 0) % 360
                if rot not in PDF_IMPORT_SUPPORTED_ROTATIONS:
                    raise UnsupportedRotationError(
                        f"Strana {page_no} má rotaci {rot}°. Podporovány jsou pouze "
                        f"rotace 0° a 180°. Soubor {os.path.basename(job.src_path)} "
                        f"nebyl OCRován, aby nevznikla falešně umístěná vrstva.")
            # Snapshot puvodnich textu pro validaci (textove strany).
            original_texts: dict[int, str] = {}
            for i in range(len(doc)):
                if (i + 1) not in job.pages_to_ocr:
                    try:
                        original_texts[i + 1] = doc[i].get_text("text") or ""
                    except Exception:
                        original_texts[i + 1] = ""
            total = len(job.pages_to_ocr)
            # Render potrebnych stran (po jedne - nizka pamet).
            images: list[Image.Image] = []
            for k, page_no in enumerate(job.pages_to_ocr):
                if is_cancelled is not None and is_cancelled():
                    raise CancelledError("Zrušeno uživatelem.")
                if on_page is not None:
                    on_page(f"strana {k + 1} z {total} (render)")
                images.append(render_page_to_image(doc[page_no - 1]))
            page_count = len(doc)
            # Volna reference na doc pred OCR - dokument zustane otevreny
            # pro overlay; fitz instance se pouziva jen z tohoto vlakna.
            def _ocr_progress(done: int, _sub: int) -> None:
                if on_page is not None:
                    on_page(f"strana {done} z {total} (OCR)")
            ocr_pages = ocr_images_sync(
                images, job.engine, job.lang_raw, job.diacritics_enabled,
                progress_cb=_ocr_progress, is_cancelled=is_cancelled)
            ocr_by_page = {pn: res for pn, res in zip(job.pages_to_ocr, ocr_pages)}
            # Prazdne OCR = chyba, ne tiche "hotovo".
            if not any(r for page in ocr_pages for r in page
                       if (getattr(r, "display_text", None) or getattr(r, "text", "") or "").strip()):
                raise OcrEngineError(
                    f"OCR engine {job.engine} nerozpoznal žádný text. "
                    f"Výstup nebyl vytvořen.")
            def _ov_progress(done: int, _sub: int) -> None:
                if on_pages_done is not None:
                    on_pages_done(done)
                if on_page is not None:
                    on_page(f"strana {done} z {total} (zápis)")
            stats = overlay_ocr_layer(
                job.src_path, ocr_by_page, PDF_IMPORT_RENDER_DPI, job.out_path,
                progress_cb=_ov_progress, is_cancelled=is_cancelled)
            logger.info("Import OK: %s -> %s stats=%s", job.src_path, job.out_path, stats)
            problems = validate_imported_pdf(job.out_path, page_count, ocr_by_page, original_texts)
            if problems:
                try:
                    if os.path.isfile(job.out_path):
                        os.remove(job.out_path)
                except OSError:
                    pass
                raise ValidationError("Validace výstupu selhala: " + "; ".join(problems))
            return JobResult(src_path=job.src_path, out_path=job.out_path,
                             status="ok", ocr_pages=list(job.pages_to_ocr),
                             message=f"OCR dokončeno. Výstup uložen jako {os.path.basename(job.out_path)}.")
        finally:
            if doc is not None:
                try:
                    doc.close()
                except Exception:
                    pass
    except CancelledError as e:
        return JobResult(src_path=job.src_path, out_path=job.out_path,
                         status="cancelled", message=str(e))
    except PdfImportError as e:
        return JobResult(src_path=job.src_path, out_path=job.out_path,
                         status="failed", message=str(e))
    except Exception as e:
        logger.exception("Neocekavana chyba importu %s", job.src_path)
        return JobResult(src_path=job.src_path, out_path=job.out_path,
                         status="failed", message=f"Neočekávaná chyba: {e}")


# ---------------------------------------------------------------------------
# Batch worker (QThread) - zadne GUI operace, pouze signaly
# ---------------------------------------------------------------------------

class PdfImportWorker(QThread):
    """Sekvencni davkovy worker: jeden soubor po druhem, jeden dokument
    otevreny v jeden okamzik. GUI ridici pres signaly."""

    job_started = pyqtSignal(str, int, int)   # src_path, poradi_od_1, celkem
    job_message = pyqtSignal(str, str)        # src_path, stavovy text stranky
    job_finished = pyqtSignal(object)         # JobResult
    overall_progress = pyqtSignal(int, int)   # hotove_strany, celkem_stran
    batch_finished = pyqtSignal(list)          # list[JobResult]

    def __init__(self, jobs: list[ImportJob]) -> None:
        super().__init__()
        self.jobs = list(jobs)
        self.results: list[JobResult] = []

    def run(self) -> None:
        total_pages = sum(len(j.pages_to_ocr) for j in self.jobs)
        done_pages = 0
        self.results = []
        for idx, job in enumerate(self.jobs):
            if self.isInterruptionRequested():
                self.results.append(JobResult(
                    src_path=job.src_path, out_path=job.out_path,
                    status="cancelled", message="Zrušeno uživatelem."))
                continue
            self.job_started.emit(job.src_path, idx + 1, len(self.jobs))

            def _on_page(text: str, _job: ImportJob = job) -> None:
                self.job_message.emit(_job.src_path, text)

            before = done_pages

            def _on_pages_done(n: int) -> None:
                self.overall_progress.emit(before + n, total_pages)

            result = process_single_job(
                job, on_page=_on_page, on_pages_done=_on_pages_done,
                is_cancelled=self.isInterruptionRequested)
            if result.status == "ok":
                done_pages = before + len(job.pages_to_ocr)
                self.overall_progress.emit(done_pages, total_pages)
            else:
                self.overall_progress.emit(done_pages, total_pages)
            self.results.append(result)
            self.job_finished.emit(result)
            if result.status == "cancelled":
                # Po zruseni jiz nic dalsiho nespoustet.
                for rest in self.jobs[idx + 1:]:
                    skipped = JobResult(src_path=rest.src_path, out_path=rest.out_path,
                                        status="cancelled", message="Zrušeno uživatelem.")
                    self.results.append(skipped)
                    self.job_finished.emit(skipped)
                break
        self.batch_finished.emit(list(self.results))
