from __future__ import annotations
import sys
import os
import tempfile
from typing import Optional, Callable

from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QComboBox, QMessageBox, QFileDialog, QDialog, QListWidget,
    QProgressDialog, QTextEdit, QCheckBox, QAbstractItemView, QGroupBox,
    QInputDialog, QScrollArea, QRadioButton, QDialogButtonBox, QButtonGroup
)
from PyQt6.QtGui import QPixmap, QShortcut, QKeySequence
from PyQt6.QtCore import Qt, QEventLoop, QSettings, QTimer, QStandardPaths
import re

# QAccessible není v PyQt6 6.11+ exponován v Python API (ověřeno
# ImportError: cannot import name 'QAccessible' from 'PyQt6.QtGui').
# Qt interně posílá Focus event automaticky při setFocus(), takže
# explicitní QAccessible.updateAccessibility() není nutné. Pokud by
# budoucí verze PyQt6 QAccessible zpřístupnila, lze jej volat
# podmíněně přes importlib – viz ScanApp._set_initial_focus().
from PIL import Image, ImageQt, ImageFilter, ImageOps
from docx import Document
import fitz
from accessible_output2.outputs.auto import Auto

from scanner_engine import NAPS2Scanner, ScanThread
from ocr_engine import create_ocr_thread, OcrResult, LANG_MAP_EASYOCR
from ocr_engine import (
    EasyOCRPreloadThread,
    is_easyocr_ready,
    is_easyocr_preparing,
    is_easyocr_failed,
    get_easyocr_status,
    _is_easyocr_model_cached,
)
from macro import Macro, MacroManager, PipelineRunner
from macro_editor import MacroEditorDialog
import pdf_import
from pdf_import import PdfImportWorker

import logging
logger = logging.getLogger(__name__)

# ------------------ Hlasový výstup ------------------
_speaker = Auto()

def speak(text: str) -> None:
    _speaker.output(text)

# ------------------ Předzpracování obrazu ------------------
def preprocess_image(img: Image.Image) -> Image.Image:
    if img.mode != "L":
        img = img.convert("L")
    img = img.filter(ImageFilter.SHARPEN)
    img = ImageOps.autocontrast(img, cutoff=2)
    return img


def alt_preprocess_image(img: Image.Image) -> Image.Image:
    if img.mode != "L":
        img = img.convert("L")
    img = img.filter(ImageFilter.MedianFilter(3))
    img = ImageOps.autocontrast(img, cutoff=1)
    return img

# ------------------ OCR PDF: neviditelná textová vrstva ------------------
# Princip: obraz skenu + neviditelný Unicode text (PDF Tr 3 / render_mode=3).
# Nikdy nepoužívat fontsize=0 ani fill_opacity=0 jako "zneviditelnění".
_OCR_PDF_FONTNAME = "ocr-cs-sans"
_OCR_PDF_BUNDLED_FONT = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "assets", "fonts", "DejaVuSans.ttf",
)
_OCR_PDF_MIN_FONTSIZE = 4.0
_OCR_PDF_MAX_START_FONTSIZE = 12.0
_OCR_PDF_SHRINK_FACTOR = 0.9


def _resolve_ocr_pdf_fontfile() -> Optional[str]:
    """Najde Unicode TTF pro OCR vrstvu. Bundlovaný má přednost před systémovým."""
    candidates = [
        _OCR_PDF_BUNDLED_FONT,
        r"C:\Windows\Fonts\DejaVuSans.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/dejavu/DejaVuSans.ttf",
    ]
    for path in candidates:
        try:
            if path and os.path.isfile(path):
                return path
        except Exception:
            continue
    return None


def _ocr_pdf_initial_fontsize(rect: "fitz.Rect") -> float:
    """Výchozí velikost písma z výšky bboxu v bodech. Nikdy <= 0."""
    try:
        h = float(rect.height)
    except Exception:
        h = 0.0
    if h <= 0:
        return 10.0
    # Odstavcové bboxy (EasyOCR paragraph=True) mohou mít výšku celé stránky –
    # start clampneme na čitelný strop, shrink-loop zbytek dořeší.
    start = h * 0.9 if h <= 24.0 else _OCR_PDF_MAX_START_FONTSIZE
    return max(_OCR_PDF_MIN_FONTSIZE, min(start, 72.0))


def _insert_ocr_textbox_invisible(page: "fitz.Page", rect: "fitz.Rect",
                                  text: str, fontname: str,
                                  start_fontsize: float) -> str:
    """Vloží neviditelný text (render_mode=3). Vrací 'ok' | 'shrunk' | 'fallback' | 'failed'.

    Nikdy nepoužije fontsize <= 0. Záporný návrat insert_textbox() řeší
    postupným zmenšováním; krajní fallback je insert_text na levý horní roh
    bboxu – text je zachován, stále neviditelný, zalogovaný.
    """
    if not text or not text.strip():
        return "ok"
    fs = max(_OCR_PDF_MIN_FONTSIZE, float(start_fontsize or 10.0))
    if fs <= 0:
        fs = 10.0
    shrunk = False
    for _ in range(25):
        try:
            ret = page.insert_textbox(rect, text, fontsize=fs,
                                      fontname=fontname, render_mode=3)
        except Exception as e:
            logger.warning("OCR PDF insert_textbox selhal (fs=%.2f): %s", fs, e)
            ret = -1.0
        if ret is not None and ret >= 0:
            return "shrunk" if shrunk else "ok"
        # Overflow – zmenšit, ale nikdy na 0.
        next_fs = max(_OCR_PDF_MIN_FONTSIZE, fs * _OCR_PDF_SHRINK_FACTOR)
        if next_fs >= fs:
            break
        fs = next_fs
        shrunk = True
        if fs <= _OCR_PDF_MIN_FONTSIZE + 1e-9:
            # Ještě jeden pokus s minimem, pak fallback.
            try:
                ret = page.insert_textbox(rect, text, fontsize=fs,
                                          fontname=fontname, render_mode=3)
            except Exception as e:
                logger.warning("OCR PDF insert_textbox (min) selhal: %s", e)
                ret = -1.0
            if ret is not None and ret >= 0:
                return "shrunk"
            break
    # Explicitní fallback: text nesmí být tiše zahozen.
    try:
        pt = fitz.Point(rect.x0 + 1, rect.y0 + _OCR_PDF_MIN_FONTSIZE)
        page.insert_text(pt, text, fontsize=_OCR_PDF_MIN_FONTSIZE,
                         fontname=fontname, render_mode=3)
        logger.warning("OCR PDF fallback insert_text pro text %r (bbox %s)",
                       text[:60], rect)
        return "fallback"
    except Exception as e:
        logger.error("OCR PDF fallback selhal pro text %r: %s", text[:60], e)
        return "failed"


def build_searchable_pdf(images: list[Image.Image],
                          ocr_pages: list[list],
                          dpi: int,
                          output_path: str) -> dict:
    """Sestaví prohledávatelné PDF: obraz + neviditelná Unicode vrstva.

    Testovatelná bez Qt. Zachovává převod px->pt (x/dpi*72) i konzistenci
    OCR-obraz vs. PDF-obraz (volající předává stejné objekty).
    Statistiky: {pages, inserted, shrunk, fallback, failed, nobbox}.
    """
    if dpi is None or int(dpi) <= 0:
        dpi = 300
    dpi = int(dpi)
    fontfile = _resolve_ocr_pdf_fontfile()
    if fontfile is None:
        logger.error("OCR PDF: Unicode font nenalezen (assets/fonts/DejaVuSans.ttf "
                     "ani systémový DejaVuSans.ttf). PDF by mělo poškozenou diakritiku.")
        raise RuntimeError(
            "Pro PDF s rozpoznaným textem chybí Unicode font "
            "(assets/fonts/DejaVuSans.ttf). Přeinstalujte aplikaci."
        )
    stats = {"pages": 0, "inserted": 0, "shrunk": 0,
             "fallback": 0, "failed": 0, "nobbox": 0}
    doc = fitz.open()
    try:
        for i, img in enumerate(images):
            page = doc.new_page(width=img.width / dpi * 72,
                                height=img.height / dpi * 72)
            tmp_path = None
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
                    tmp_path = tmp.name
                    img.save(tmp_path, "JPEG", quality=95)
                page.insert_image(page.rect, filename=tmp_path)
            finally:
                if tmp_path and os.path.exists(tmp_path):
                    try:
                        os.remove(tmp_path)
                    except OSError:
                        pass
            try:
                page.insert_font(fontname=_OCR_PDF_FONTNAME, fontfile=fontfile)
            except Exception as e:
                logger.error("OCR PDF: registrace fontu selhala: %s", e)
                raise RuntimeError(f"Registrace Unicode fontu selhala: {e}")
            stats["pages"] += 1
            nobbox_texts: list[str] = []
            if i < len(ocr_pages):
                for item in ocr_pages[i]:
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
                    if bbox is None:
                        nobbox_texts.append(txt)
                        continue
                    try:
                        x0, y0 = float(bbox[0][0]), float(bbox[0][1])
                        x1, y1 = float(bbox[2][0]), float(bbox[2][1])
                    except Exception as e:
                        logger.warning("OCR PDF: neplatný bbox %r (%s) – použit fallback.", bbox, e)
                        nobbox_texts.append(txt)
                        continue
                    if x1 <= x0 or y1 <= y0:
                        logger.warning("OCR PDF: degenerovaný bbox %r – použit fallback.", bbox)
                        nobbox_texts.append(txt)
                        continue
                    rect = fitz.Rect(x0 / dpi * 72, y0 / dpi * 72,
                                     x1 / dpi * 72, y1 / dpi * 72)
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
                logger.warning("OCR PDF: strana %d má %d položek bez bbox – "
                               "fallback do spodního pruhu.", i + 1, len(nobbox_texts))
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
        doc.save(output_path)
    finally:
        try:
            doc.close()
        except Exception:
            pass
    if stats["failed"]:
        logger.error("OCR PDF: %d textů se nepodařilo vložit: %s", stats["failed"], output_path)
    return stats

# ------------------ Náhled dialog ------------------
class PreviewDialog(QDialog):
    def __init__(self, pil_img: Image.Image, page_num: int = 1) -> None:
        super().__init__()
        self.setWindowTitle("Náhled skenu")
        self.image = pil_img
        self.page_num = page_num

        self.label = QLabel()
        self.update_image()

        layout = QVBoxLayout()
        layout.addWidget(self.label)

        btn_layout = QHBoxLayout()
        btn_rotate = QPushButton("Otočit")
        btn_rotate.setAccessibleName("Otočit obrázek")
        btn_rotate.clicked.connect(self.rotate_image)
        btn_layout.addWidget(btn_rotate)

        btn_ok = QPushButton("Potvrdit náhled")
        btn_ok.setAccessibleName("Potvrdit náhled")
        btn_ok.clicked.connect(self.accept)
        btn_layout.addWidget(btn_ok)

        layout.addLayout(btn_layout)
        self.setLayout(layout)

        QTimer.singleShot(300, lambda: speak(
            f"Náhled stránky {page_num}. Otočte obrázek nebo potvrďte náhled."
        ))

    def update_image(self) -> None:
        qimg = ImageQt.ImageQt(self.image)
        pix = QPixmap.fromImage(qimg).scaled(
            500, 700, Qt.AspectRatioMode.KeepAspectRatio
        )
        self.label.setPixmap(pix)

    def rotate_image(self) -> None:
        self.image = self.image.rotate(-90, expand=True)
        self.update_image()
        speak("Obrázek otočen.")

# ------------------ Dialog pro editaci OCR výsledku ------------------
class OcrPreviewDialog(QDialog):
    RESULT_RETRY = 2  # vlastní kód pro „Zkusit jiný engine“

    def __init__(self, text: str, parent: QWidget | None = None, alt_engine_name: str = "") -> None:
        super().__init__(parent)
        self.setWindowTitle("Náhled OCR textu")
        self.setMinimumSize(600, 400)
        self._result_action = "reject"  # accept | retry | reject
        self._alt_engine_name = alt_engine_name

        self.text_edit = QTextEdit()
        self.text_edit.setPlainText(text)
        self.text_edit.setAccessibleName("OCR text k editaci")

        btn_layout = QHBoxLayout()
        btn_read = QPushButton("Přečíst text")
        btn_read.setAccessibleName("Přečíst text hlasem")
        btn_read.clicked.connect(self.read_text)
        btn_layout.addWidget(btn_read)

        btn_save = QPushButton("Použít aktuální výsledek")
        btn_save.setAccessibleName("Použít aktuální výsledek a pokračovat")
        btn_save.setDefault(True)
        btn_save.clicked.connect(self._on_accept)
        btn_layout.addWidget(btn_save)
        self._btn_save = btn_save

        self._btn_retry: QPushButton | None = None
        if alt_engine_name:
            btn_retry = QPushButton("Zkusit jiný OCR engine")
            # Jednorázový pokus, nemění engine_combo
            btn_retry.setAccessibleName(f"Zkusit jiný OCR engine – {alt_engine_name}")
            btn_retry.setAccessibleDescription("Spustí jednorázové OCR druhým enginem. Původní výsledek zůstane zachován do úspěchu nového pokusu.")
            btn_retry.clicked.connect(self._on_retry)
            btn_layout.addWidget(btn_retry)
            self._btn_retry = btn_retry

        btn_cancel = QPushButton("Zrušit")
        btn_cancel.setAccessibleName("Zrušit")
        btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(btn_cancel)
        self._btn_cancel = btn_cancel

        layout = QVBoxLayout()
        layout.addWidget(self.text_edit)
        layout.addLayout(btn_layout)
        self.setLayout(layout)

        # Tab order: text_edit -> Přečíst -> Použít -> Zkusit jiný -> Zrušit
        self.setTabOrder(self.text_edit, btn_read)
        self.setTabOrder(btn_read, btn_save)
        if self._btn_retry is not None:
            self.setTabOrder(btn_save, self._btn_retry)
            self.setTabOrder(self._btn_retry, btn_cancel)
        else:
            self.setTabOrder(btn_save, btn_cancel)

    def _on_accept(self) -> None:
        self._result_action = "accept"
        self.accept()

    def _on_retry(self) -> None:
        self._result_action = "retry"
        # Ukončit dialog s vlastním kódem
        self.done(OcrPreviewDialog.RESULT_RETRY)

    def result_action(self) -> str:
        return self._result_action

    def read_text(self) -> None:
        speak(self.text_edit.toPlainText())

    def showEvent(self, event) -> None:
        super().showEvent(event)
        self.text_edit.setFocus()
        if self._alt_engine_name:
            speak(f"Náhled OCR dokončen. Můžete použít aktuální výsledek nebo zkusit jiný OCR engine {self._alt_engine_name}.")

    def get_text(self) -> str:
        return self.text_edit.toPlainText()


# ------------------ Pomocné funkce pro stránkování ------------------
_PAGE_HEADER_RE = re.compile(r"^---\s*Stránka\s+(\d+)\s*---\s*$", re.MULTILINE)


def _split_text_by_page_headers(text: str) -> list[str] | None:
    """Rozdělí text podle headerů '--- Stránka N ---' na list per-page.

    Vrátí None pokud text neobsahuje žádné headery nebo je neplatný.
    """
    if not text or "--- Stránka" not in text:
        return None
    # Najdi všechny headery
    matches = list(_PAGE_HEADER_RE.finditer(text))
    if not matches:
        return None
    pages: list[str] = []
    for idx, m in enumerate(matches):
        start = m.end()
        end = matches[idx + 1].start() if idx + 1 < len(matches) else len(text)
        page_content = text[start:end].strip("\n")
        # Odstraň přebytečné okolní nové řádky, ale zachovej vnitřní
        # strip jen koncové \n, pak strip surrounding whitespace per page
        pages.append(page_content.strip())
    return pages


# ------------------ Dialog pro výběr formátu uložení (legacy) ------------------
class SaveDocumentDialog(QDialog):
    """Přístupný dialog 'Uložit dokument' – výběr formátu s lidskými názvy (ponechán pro kompatibilitu makro)."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Uložit dokument")
        self.setMinimumWidth(520)
        self._selected_format: str = "txt"

        layout = QVBoxLayout(self)

        lbl_info = QLabel("Vyberte formát, ve kterém chcete dokument uložit.")
        lbl_info.setWordWrap(True)
        lbl_info.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        layout.addWidget(lbl_info)

        # Radio buttons s lidskými názvy a popisem
        self.rb_txt = QRadioButton("Textový dokument (.txt)")
        self.rb_txt.setAccessibleName("Textový dokument (.txt)")
        self.rb_txt.setAccessibleDescription("Pouhý text bez zachování vzhledu stránky.")
        self.rb_txt.setChecked(True)
        layout.addWidget(self.rb_txt)
        lbl_txt_desc = QLabel("Pouhý text bez zachování vzhledu stránky.")
        lbl_txt_desc.setWordWrap(True)
        lbl_txt_desc.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        lbl_txt_desc.setStyleSheet("color: palette(mid); margin-left: 22px;")
        layout.addWidget(lbl_txt_desc)

        self.rb_docx = QRadioButton("Word dokument (.docx)")
        self.rb_docx.setAccessibleName("Word dokument (.docx)")
        self.rb_docx.setAccessibleDescription("Upravitelný dokument pro Microsoft Word a další kompatibilní programy.")
        layout.addWidget(self.rb_docx)
        lbl_docx_desc = QLabel("Upravitelný dokument pro Microsoft Word a další kompatibilní programy.")
        lbl_docx_desc.setWordWrap(True)
        lbl_docx_desc.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        lbl_docx_desc.setStyleSheet("color: palette(mid); margin-left: 22px;")
        layout.addWidget(lbl_docx_desc)

        self.rb_pdf = QRadioButton("PDF s OCR (.pdf)")
        self.rb_pdf.setAccessibleName("PDF s OCR (.pdf)")
        self.rb_pdf.setAccessibleDescription("Dokument se zachovaným vzhledem stránek a skrytou textovou vrstvou pro vyhledávání a odečítání textu.")
        layout.addWidget(self.rb_pdf)
        lbl_pdf_desc = QLabel("Dokument se zachovaným vzhledem stránek a skrytou textovou vrstvou pro vyhledávání a odečítání textu.")
        lbl_pdf_desc.setWordWrap(True)
        lbl_pdf_desc.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        lbl_pdf_desc.setStyleSheet("color: palette(mid); margin-left: 22px;")
        layout.addWidget(lbl_pdf_desc)

        # Button group pro logické propojení
        self._group = QButtonGroup(self)
        self._group.addButton(self.rb_txt, 0)
        self._group.addButton(self.rb_docx, 1)
        self._group.addButton(self.rb_pdf, 2)

        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, self)
        btn_ok = btn_box.button(QDialogButtonBox.StandardButton.Ok)
        btn_ok.setText("Uložit")
        btn_ok.setAccessibleName("Uložit v zvoleném formátu")
        btn_ok.setDefault(True)
        btn_cancel = btn_box.button(QDialogButtonBox.StandardButton.Cancel)
        btn_cancel.setText("Zrušit")
        btn_cancel.setAccessibleName("Zrušit ukládání")
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

        # Tab order je přirozený: rb_txt -> rb_docx -> rb_pdf -> Ok -> Cancel
        self.setTabOrder(self.rb_txt, self.rb_docx)
        self.setTabOrder(self.rb_docx, self.rb_pdf)
        self.setTabOrder(self.rb_pdf, btn_ok)
        self.setTabOrder(btn_ok, btn_cancel)

    def showEvent(self, event) -> None:
        super().showEvent(event)
        # Počáteční fokus na první volbu
        self.rb_txt.setFocus(Qt.FocusReason.OtherFocusReason)
        speak("Dialog Uložit dokument. Vyberte formát, ve kterém chcete dokument uložit.")

    def selected_format(self) -> str:
        if self.rb_docx.isChecked():
            return "docx"
        if self.rb_pdf.isChecked():
            return "pdf"
        return "txt"


# ------------------ Dvoustupňové dialogy pro oddělené ukládání ------------------
class ChooseSaveKindDialog(QDialog):
    """První krok 'Co chcete uložit?' – odděluje uložení bez OCR a s OCR."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Co chcete uložit?")
        self.setMinimumWidth(520)
        layout = QVBoxLayout(self)
        lbl_info = QLabel("Co chcete uložit?")
        lbl_info.setWordWrap(True)
        lbl_info.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        # accessible name for NVDA heading
        lbl_info.setAccessibleName("Co chcete uložit")
        layout.addWidget(lbl_info)

        self.rb_images = QRadioButton("Naskenované stránky")
        self.rb_images.setAccessibleName("Naskenované stránky")
        self.rb_images.setAccessibleDescription("Uloží stránky bez OCR jako PDF. Uloží stránky jako obrázky. OCR nebude proveden.")
        self.rb_images.setChecked(True)
        layout.addWidget(self.rb_images)
        lbl_images_desc = QLabel("Uloží stránky bez OCR jako PDF.")
        lbl_images_desc.setWordWrap(True)
        lbl_images_desc.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        lbl_images_desc.setStyleSheet("color: palette(mid); margin-left: 22px;")
        layout.addWidget(lbl_images_desc)
        lbl_images_detail = QLabel("Uloží stránky jako obrázky. OCR nebude proveden.")
        lbl_images_detail.setWordWrap(True)
        lbl_images_detail.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        lbl_images_detail.setStyleSheet("color: palette(mid); margin-left: 22px; font-style: italic;")
        layout.addWidget(lbl_images_detail)

        self.rb_ocr = QRadioButton("OCR dokument")
        self.rb_ocr.setAccessibleName("OCR dokument")
        self.rb_ocr.setAccessibleDescription("Uloží již rozpoznaný text jako TXT, DOCX nebo PDF s OCR. Vyžaduje předchozí OCR.")
        layout.addWidget(self.rb_ocr)
        lbl_ocr_desc = QLabel("Uloží již rozpoznaný text jako TXT, DOCX nebo PDF s OCR.")
        lbl_ocr_desc.setWordWrap(True)
        lbl_ocr_desc.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        lbl_ocr_desc.setStyleSheet("color: palette(mid); margin-left: 22px;")
        layout.addWidget(lbl_ocr_desc)

        self._group = QButtonGroup(self)
        self._group.addButton(self.rb_images, 0)
        self._group.addButton(self.rb_ocr, 1)

        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, self)
        btn_ok = btn_box.button(QDialogButtonBox.StandardButton.Ok)
        btn_ok.setText("Pokračovat")
        btn_ok.setAccessibleName("Pokračovat v ukládání")
        btn_ok.setDefault(True)
        btn_cancel = btn_box.button(QDialogButtonBox.StandardButton.Cancel)
        btn_cancel.setText("Zrušit")
        btn_cancel.setAccessibleName("Zrušit ukládání")
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)
        self.setTabOrder(self.rb_images, self.rb_ocr)
        self.setTabOrder(self.rb_ocr, btn_ok)
        self.setTabOrder(btn_ok, btn_cancel)
        self._btn_ok = btn_ok
        self._btn_cancel = btn_cancel

    def showEvent(self, event) -> None:
        super().showEvent(event)
        self.rb_images.setFocus(Qt.FocusReason.OtherFocusReason)
        speak("Dialog Co chcete uložit. Vyberte Naskenované stránky bez OCR nebo OCR dokument s rozpoznaným textem.")

    def selected_kind(self) -> str:
        if self.rb_ocr.isChecked():
            return "ocr"
        return "images"


class ChooseOcrFormatDialog(QDialog):
    """Druhý krok – výběr konkrétního OCR formátu."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Uložit OCR dokument")
        self.setMinimumWidth(520)
        layout = QVBoxLayout(self)
        lbl_info = QLabel("Vyberte formát OCR dokumentu.")
        lbl_info.setWordWrap(True)
        lbl_info.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        layout.addWidget(lbl_info)

        self.rb_txt = QRadioButton("Textový dokument (.txt)")
        self.rb_txt.setAccessibleName("Textový dokument (.txt)")
        self.rb_txt.setAccessibleDescription("Pouhý text bez zachování vzhledu stránky. Vyžaduje OCR.")
        self.rb_txt.setChecked(True)
        layout.addWidget(self.rb_txt)
        lbl_txt_desc = QLabel("Pouhý text bez zachování vzhledu stránky.")
        lbl_txt_desc.setWordWrap(True)
        lbl_txt_desc.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        lbl_txt_desc.setStyleSheet("color: palette(mid); margin-left: 22px;")
        layout.addWidget(lbl_txt_desc)

        self.rb_docx = QRadioButton("Word dokument (.docx)")
        self.rb_docx.setAccessibleName("Word dokument (.docx)")
        self.rb_docx.setAccessibleDescription("Upravitelný dokument pro Microsoft Word. Vyžaduje OCR.")
        layout.addWidget(self.rb_docx)
        lbl_docx_desc = QLabel("Upravitelný dokument pro Microsoft Word a další kompatibilní programy.")
        lbl_docx_desc.setWordWrap(True)
        lbl_docx_desc.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        lbl_docx_desc.setStyleSheet("color: palette(mid); margin-left: 22px;")
        layout.addWidget(lbl_docx_desc)

        self.rb_pdf_ocr = QRadioButton("PDF s rozpoznaným textem (.pdf)")
        self.rb_pdf_ocr.setAccessibleName("PDF s rozpoznaným textem (.pdf)")
        self.rb_pdf_ocr.setAccessibleDescription("Zachová vzhled naskenovaných stránek a přidá textovou vrstvu z OCR. Vyžaduje OCR.")
        layout.addWidget(self.rb_pdf_ocr)
        lbl_pdf_desc = QLabel("Zachová vzhled naskenovaných stránek a přidá textovou vrstvu z OCR.")
        lbl_pdf_desc.setWordWrap(True)
        lbl_pdf_desc.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        lbl_pdf_desc.setStyleSheet("color: palette(mid); margin-left: 22px;")
        layout.addWidget(lbl_pdf_desc)

        self._group = QButtonGroup(self)
        self._group.addButton(self.rb_txt, 0)
        self._group.addButton(self.rb_docx, 1)
        self._group.addButton(self.rb_pdf_ocr, 2)

        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, self)
        btn_ok = btn_box.button(QDialogButtonBox.StandardButton.Ok)
        btn_ok.setText("Uložit")
        btn_ok.setAccessibleName("Uložit v zvoleném OCR formátu")
        btn_ok.setDefault(True)
        btn_cancel = btn_box.button(QDialogButtonBox.StandardButton.Cancel)
        btn_cancel.setText("Zrušit")
        btn_cancel.setAccessibleName("Zrušit ukládání")
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)
        self.setTabOrder(self.rb_txt, self.rb_docx)
        self.setTabOrder(self.rb_docx, self.rb_pdf_ocr)
        self.setTabOrder(self.rb_pdf_ocr, btn_ok)
        self.setTabOrder(btn_ok, btn_cancel)
        self._btn_ok = btn_ok
        self._btn_cancel = btn_cancel

    def showEvent(self, event) -> None:
        super().showEvent(event)
        self.rb_txt.setFocus(Qt.FocusReason.OtherFocusReason)
        speak("Dialog Uložit OCR dokument. Vyberte formát TXT, DOCX nebo PDF s rozpoznaným textem.")

    def selected_format(self) -> str:
        if self.rb_docx.isChecked():
            return "docx"
        if self.rb_pdf_ocr.isChecked():
            return "pdf_ocr"
        return "txt"


# ------------------ Dialogy pro import PDF ------------------
class PdfImportSummaryDialog(QDialog):
    """Souhrn analýzy importovaných PDF – přístupný, NVDA-first.

    Zobrazí per-file výsledek (počet stran, strany potřebující OCR).
    Tlačítka: Pokračovat k OCR / Zrušit.
    """

    def __init__(self, summary_text: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Import PDF – souhrn analýzy")
        self.setMinimumWidth(560)
        layout = QVBoxLayout(self)
        lbl = QLabel(summary_text)
        lbl.setWordWrap(True)
        lbl.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        lbl.setAccessibleName("Souhrn analýzy PDF")
        lbl.setAccessibleDescription(summary_text)
        layout.addWidget(lbl)
        self._summary_text = summary_text

        btn_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, self)
        btn_ok = btn_box.button(QDialogButtonBox.StandardButton.Ok)
        btn_ok.setText("Pokračovat k OCR")
        btn_ok.setAccessibleName("Pokračovat k OCR")
        btn_ok.setAccessibleDescription(
            "Pokračuje k výběru OCR enginu a jazyka. OCR proběhne pouze na stránkách bez textu.")
        btn_ok.setDefault(True)
        btn_cancel = btn_box.button(QDialogButtonBox.StandardButton.Cancel)
        btn_cancel.setText("Zrušit")
        btn_cancel.setAccessibleName("Zrušit import")
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)
        self._btn_ok = btn_ok
        self.setTabOrder(btn_ok, btn_cancel)

    def showEvent(self, event) -> None:
        super().showEvent(event)
        self._btn_ok.setFocus(Qt.FocusReason.OtherFocusReason)
        speak("Souhrn analýzy PDF. " + self._summary_text)


class ImportOcrSettingsDialog(QDialog):
    """OCR nastavení pro import – engine + jazyk.

    Výchozí hodnoty přebírá z hlavní aplikace (engine_combo / lang_combo),
    po potvrzení je volající propíše zpět – žádný druhý systém nastavení.
    """

    LANG_ITEMS = [
        "ces (Čeština)", "eng (Angličtina)", "deu (Němčina)",
        "fra (Francouzština)", "ita (Italština)", "pol (Polština)",
    ]

    def __init__(self, engine_default: str, lang_index_default: int,
                 info_text: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("OCR nastavení pro import")
        self.setMinimumWidth(480)
        layout = QVBoxLayout(self)

        lbl_info = QLabel(info_text)
        lbl_info.setWordWrap(True)
        lbl_info.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        lbl_info.setAccessibleName("Informace o rozsahu OCR")
        layout.addWidget(lbl_info)

        lbl_engine = QLabel("OCR engine:")
        self.engine_combo = QComboBox()
        self.engine_combo.setAccessibleName("OCR engine pro import")
        self.engine_combo.setAccessibleDescription("Vyberte Tesseract nebo EasyOCR pro rozpoznání stránek bez textu.")
        self.engine_combo.addItems(["Tesseract", "EasyOCR"])
        idx = self.engine_combo.findText(engine_default)
        if idx >= 0:
            self.engine_combo.setCurrentIndex(idx)
        lbl_engine.setBuddy(self.engine_combo)
        layout.addWidget(lbl_engine)
        layout.addWidget(self.engine_combo)

        lbl_lang = QLabel("Jazyk OCR:")
        self.lang_combo = QComboBox()
        self.lang_combo.setAccessibleName("Jazyk OCR pro import")
        self.lang_combo.setAccessibleDescription("Vyberte jazyk dokumentu pro rozpoznání textu.")
        self.lang_combo.addItems(list(self.LANG_ITEMS))
        if 0 <= lang_index_default < self.lang_combo.count():
            self.lang_combo.setCurrentIndex(lang_index_default)
        lbl_lang.setBuddy(self.lang_combo)
        layout.addWidget(lbl_lang)
        layout.addWidget(self.lang_combo)

        btn_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, self)
        btn_ok = btn_box.button(QDialogButtonBox.StandardButton.Ok)
        btn_ok.setText("Spustit OCR")
        btn_ok.setAccessibleName("Spustit OCR importovaných PDF")
        btn_ok.setDefault(True)
        btn_cancel = btn_box.button(QDialogButtonBox.StandardButton.Cancel)
        btn_cancel.setText("Zrušit")
        btn_cancel.setAccessibleName("Zrušit import")
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)
        self.setTabOrder(self.engine_combo, self.lang_combo)
        self.setTabOrder(self.lang_combo, btn_ok)
        self.setTabOrder(btn_ok, btn_cancel)
        self._btn_ok = btn_ok

    def showEvent(self, event) -> None:
        super().showEvent(event)
        self.engine_combo.setFocus(Qt.FocusReason.OtherFocusReason)
        speak("Dialog OCR nastavení pro import. Vyberte engine a jazyk a potvrďte Spustit OCR.")

    def selected_engine(self) -> str:
        return self.engine_combo.currentText()

    def selected_lang_index(self) -> int:
        return self.lang_combo.currentIndex()

    def selected_lang_text(self) -> str:
        return self.lang_combo.currentText()


# ------------------ Hlavní aplikace ------------------
class ScanApp(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Hlasový skener – NAPS2 + OCR")
        self.setMinimumSize(500, 600)

        self.settings = QSettings("HlasovySkener", "HlasovySkener")
        self.scanner = NAPS2Scanner()
        self.scanned_images: list[Image.Image] = []
        self.raw_scanned_images: list[Image.Image] = []
        self.last_ocr_results: list[list[OcrResult]] = []
        self.last_scanned_image: Optional[bool] = None
        self._scan_cancelled = False
        self._ocr_mode = "interactive"
        self._last_text = ""
        self._last_original_text = ""
        self._last_edited_pages: list[str] | None = None
        self._diacritics_failed = False
        self._diacritics_enabled_cache = False
        self._ocr_total_pages: int = 0
        self._ocr_last_announced_page: int = 0
        # Pro dvoustupňové ukládání – zachování volby přes OCR
        self._pending_save_format: str | None = None

        self.macro_manager = MacroManager()
        self.macro_manager.load_all()
        self._macro_runner = PipelineRunner(self)
        self._macro_shortcuts: list[QShortcut] = []
        # Flag pro jednorázové nastavení počátečního fokusu po zobrazení okna
        self._initial_focus_done: bool = False

        # EasyOCR preload – nesmí blokovat GUI
        self._easyocr_preload_thread: EasyOCRPreloadThread | None = None
        self._easyocr_pending_ocr: dict | None = None  # uložené parametry OCR během přípravy
        self._easyocr_status_text: str = ""

        self._build_ui()
        self._setup_shortcuts()
        self._load_settings()
        # Propojení změn jazyka/enginu s preloadem (mimo GUI blokaci)
        try:
            self.lang_combo.currentTextChanged.connect(self._on_engine_or_lang_changed)
            self.engine_combo.currentTextChanged.connect(self._on_engine_or_lang_changed)
        except Exception:
            pass
        # Preload naplánovat až po event loop – GUI se zobrazí → fokus → příprava na pozadí
        QTimer.singleShot(0, self._preload_easyocr_if_needed)

    def speak(self, text: str) -> None:
        speak(text)

    # ---------- UI ----------
    def _build_ui(self) -> None:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        # QScrollArea nesmí krást počáteční fokus – je to kontejner
        scroll.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        # Viewport scroll area také nesmí být focusable (Qt default může být StrongFocus u některých stylů)
        scroll.viewport().setFocusPolicy(Qt.FocusPolicy.NoFocus)

        container = QWidget()
        container.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        layout = QVBoxLayout(container)
        layout.setSpacing(12)
        layout.setContentsMargins(15, 15, 15, 15)

        # -- Nastavení OCR --
        ocr_group = QGroupBox("Nastavení OCR")
        ocr_layout = QVBoxLayout()

        lbl_lang = QLabel("Jazyk OCR:")
        self.lang_combo = QComboBox()
        self.lang_combo.setAccessibleName("Jazyk OCR")
        self.lang_combo.addItems([
            "ces (Čeština)", "eng (Angličtina)", "deu (Němčina)",
            "fra (Francouzština)", "ita (Italština)", "pol (Polština)"
        ])
        lbl_lang.setBuddy(self.lang_combo)
        ocr_layout.addWidget(lbl_lang)
        ocr_layout.addWidget(self.lang_combo)

        lbl_engine = QLabel("OCR engine:")
        self.engine_combo = QComboBox()
        self.engine_combo.addItems(["Tesseract", "EasyOCR"])
        self.engine_combo.setAccessibleName("OCR engine")
        lbl_engine.setBuddy(self.engine_combo)
        ocr_layout.addWidget(lbl_engine)
        ocr_layout.addWidget(self.engine_combo)

        self.preprocess_cb = QCheckBox("Předzpracovat obraz (doostřit, kontrast)")
        self.preprocess_cb.setAccessibleName("Předzpracování obrazu")
        self.preprocess_cb.setChecked(True)
        ocr_layout.addWidget(self.preprocess_cb)

        lbl_diac = QLabel("Oprava české diakritiky:")
        self.diacritics_cb = QCheckBox("Oprava české diakritiky")
        self.diacritics_cb.setAccessibleName("Oprava české diakritiky")
        self.diacritics_cb.setChecked(False)
        self.diacritics_cb.toggled.connect(self._on_diacritics_toggled)
        lbl_diac.setBuddy(self.diacritics_cb)
        ocr_layout.addWidget(lbl_diac)
        ocr_layout.addWidget(self.diacritics_cb)

        # Stav EasyOCR – přístupný text + retry akce
        self.easyocr_status_label = QLabel("EasyOCR: nepřipraven")
        self.easyocr_status_label.setAccessibleName("Stav EasyOCR")
        self.easyocr_status_label.setAccessibleDescription("Informuje o stavu přípravy EasyOCR pro odečítání textu.")
        self.easyocr_status_label.setWordWrap(True)
        self.easyocr_status_label.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.easyocr_status_label.setStyleSheet("color: palette(mid);")
        ocr_layout.addWidget(self.easyocr_status_label)

        self.btn_retry_easyocr = QPushButton("Připravit EasyOCR znovu")
        self.btn_retry_easyocr.setAccessibleName("Připravit EasyOCR znovu")
        self.btn_retry_easyocr.setAccessibleDescription("Znovu připraví EasyOCR. Používá stejný worker a nevytvoří duplicitní stahování.")
        self.btn_retry_easyocr.clicked.connect(self._on_retry_easyocr_clicked)
        self.btn_retry_easyocr.hide()
        ocr_layout.addWidget(self.btn_retry_easyocr)

        ocr_group.setLayout(ocr_layout)
        layout.addWidget(ocr_group)

        # -- Nastavení skeneru --
        scan_group = QGroupBox("Nastavení skeneru")
        scan_layout = QVBoxLayout()

        self.batch_cb = QCheckBox("Dávkový režim (automaticky všechny stránky)")
        self.batch_cb.setAccessibleName("Dávkový režim")
        scan_layout.addWidget(self.batch_cb)

        lbl_device = QLabel("Vyber profil NAPS2:")
        self.device_combo = QComboBox()
        self.device_combo.setAccessibleName("Vyber profil NAPS2")
        lbl_device.setBuddy(self.device_combo)
        self.reload_devices()
        scan_layout.addWidget(lbl_device)
        scan_layout.addWidget(self.device_combo)

        lbl_source = QLabel("Zdroj papíru:")
        self.source_combo = QComboBox()
        self.source_combo.setAccessibleName("Zdroj papíru")
        self.source_combo.addItems(["Sklo", "Podavač"])
        lbl_source.setBuddy(self.source_combo)
        scan_layout.addWidget(lbl_source)
        scan_layout.addWidget(self.source_combo)

        lbl_dpi = QLabel("Rozlišení DPI:")
        self.dpi_combo = QComboBox()
        self.dpi_combo.setAccessibleName("Rozlišení DPI")
        self.dpi_combo.setEditable(False)
        self.dpi_combo.addItem("200 DPI – rychlé skenování", 200)
        self.dpi_combo.addItem("300 DPI – doporučeno pro OCR", 300)
        self.dpi_combo.addItem("400 DPI – menší text", 400)
        self.dpi_combo.addItem("600 DPI – velmi malý text", 600)
        self.dpi_combo.addItem("1200 DPI – vysoká kvalita", 1200)
        idx = self.dpi_combo.findData(300)
        if idx >= 0:
            self.dpi_combo.setCurrentIndex(idx)
        lbl_dpi.setBuddy(self.dpi_combo)
        scan_layout.addWidget(lbl_dpi)
        scan_layout.addWidget(self.dpi_combo)

        lbl_color = QLabel("Režim:")
        self.color_combo = QComboBox()
        self.color_combo.setAccessibleName("Režim barev")
        self.color_combo.addItems(["Barevný", "Šedý", "ČB"])
        lbl_color.setBuddy(self.color_combo)
        scan_layout.addWidget(lbl_color)
        scan_layout.addWidget(self.color_combo)

        scan_group.setLayout(scan_layout)
        layout.addWidget(scan_group)

        # -- Tlačítka skenování --
        scan_btn_layout = QHBoxLayout()
        self.btn_scan = QPushButton("Skenovat stránku (Ctrl+S)")
        self.btn_scan.setAccessibleName("Skenovat stránku")
        self.btn_scan.clicked.connect(self.scan_pages)
        scan_btn_layout.addWidget(self.btn_scan)

        self.btn_scan_next = QPushButton("Skenovat další stránku (Ctrl+Shift+N)")
        self.btn_scan_next.setAccessibleName("Skenovat další stránku – přidat ke stávajícímu dokumentu")
        self.btn_scan_next.clicked.connect(self.scan_next_page)
        scan_btn_layout.addWidget(self.btn_scan_next)

        self.btn_scan_all = QPushButton("Skenovat vše (Ctrl+Shift+S)")
        self.btn_scan_all.setAccessibleName("Skenovat všechny stránky dávkově")
        self.btn_scan_all.clicked.connect(self.scan_all_pages)
        scan_btn_layout.addWidget(self.btn_scan_all)
        layout.addLayout(scan_btn_layout)

        # -- Seznam stránek --
        self.page_list = QListWidget()
        self.page_list.setAccessibleName("Seznam naskenovaných stránek")
        self.page_list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        layout.addWidget(QLabel("Naskenované stránky:"))
        layout.addWidget(self.page_list)

        page_btn_layout = QHBoxLayout()
        btn_delete_page = QPushButton("Smazat stránku")
        btn_delete_page.setAccessibleName("Smazat vybranou stránku")
        btn_delete_page.clicked.connect(self.delete_page)
        page_btn_layout.addWidget(btn_delete_page)

        btn_clear_pages = QPushButton("Smazat vše")
        btn_clear_pages.setAccessibleName("Smazat všechny stránky")
        btn_clear_pages.clicked.connect(self.clear_pages)
        page_btn_layout.addWidget(btn_clear_pages)
        layout.addLayout(page_btn_layout)

        # -- Tlačítka OCR a uložení – logicky oddělené --
        ocr_btn_layout = QHBoxLayout()
        self.btn_ocr = QPushButton("Spustit OCR (Ctrl+O)")
        self.btn_ocr.setAccessibleName("Spustit OCR – rozpozná text ze stránek")
        self.btn_ocr.clicked.connect(self.run_ocr)
        self.btn_ocr.setEnabled(False)
        ocr_btn_layout.addWidget(self.btn_ocr)

        self.btn_save_document = QPushButton("Uložit dokument (Ctrl+U)")
        self.btn_save_document.setAccessibleName("Uložit dokument – uloží naskenované stránky nebo rozpoznaný text")
        self.btn_save_document.clicked.connect(self.save_document)
        self.btn_save_document.setEnabled(False)
        ocr_btn_layout.addWidget(self.btn_save_document)

        self.btn_read = QPushButton("Přečíst text (Ctrl+P)")
        self.btn_read.setAccessibleName("Přečíst naposledy rozpoznaný text")
        self.btn_read.clicked.connect(self.read_last_text)
        self.btn_read.setEnabled(False)
        ocr_btn_layout.addWidget(self.btn_read)
        layout.addLayout(ocr_btn_layout)

        btn_export_img = QPushButton("Exportovat obrázky (Ctrl+E)")
        btn_export_img.setAccessibleName("Exportovat naskenované obrázky")
        btn_export_img.clicked.connect(self.export_images)
        layout.addWidget(btn_export_img)

        self.btn_import_pdf = QPushButton("Importovat PDF (Ctrl+I)")
        self.btn_import_pdf.setAccessibleName("Importovat PDF")
        self.btn_import_pdf.setAccessibleDescription(
            "Vybere jeden nebo více PDF souborů, zjistí které stránky potřebují OCR "
            "a vytvoří nové prohledávatelné PDF. Původní soubor zůstane zachován."
        )
        self.btn_import_pdf.clicked.connect(self.import_pdfs)
        layout.addWidget(self.btn_import_pdf)

        # -- Uživatelská makra --
        macro_group = QGroupBox("Uživatelská makra")
        self._macro_container = QVBoxLayout()
        self._macro_buttons_layout = QVBoxLayout()
        self._macro_container.addLayout(self._macro_buttons_layout)
        btn_edit_macros = QPushButton("Spravovat makra...")
        btn_edit_macros.setAccessibleName("Otevřít editor maker")
        btn_edit_macros.clicked.connect(self._open_macro_editor)
        self._macro_container.addWidget(btn_edit_macros)
        macro_group.setLayout(self._macro_container)
        layout.addWidget(macro_group)

        scroll.setWidget(container)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.addWidget(scroll)
        self.setLayout(main_layout)
        self.setStyleSheet("""
            QPushButton, QComboBox {
                padding: 6px;
                min-height: 1.5em;
            }
            QCheckBox {
                padding: 6px;
                min-height: 1.75em;
                spacing: 8px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
            }
            QGroupBox {
                margin-top: 1.2em;
                padding-top: 8px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 4px;
            }
        """)
        self._rebuild_macro_buttons()

        # -- Tab order --
        self.setTabOrder(self.lang_combo, self.engine_combo)
        self.setTabOrder(self.engine_combo, self.btn_retry_easyocr)
        self.setTabOrder(self.btn_retry_easyocr, self.preprocess_cb)
        self.setTabOrder(self.preprocess_cb, self.diacritics_cb)
        self.setTabOrder(self.diacritics_cb, self.batch_cb)
        self.setTabOrder(self.batch_cb, self.device_combo)
        self.setTabOrder(self.device_combo, self.source_combo)
        self.setTabOrder(self.source_combo, self.dpi_combo)
        self.setTabOrder(self.dpi_combo, self.color_combo)
        self.setTabOrder(self.color_combo, self.btn_scan)
        self.setTabOrder(self.btn_scan, self.btn_scan_next)
        self.setTabOrder(self.btn_scan_next, self.btn_scan_all)
        self.setTabOrder(self.btn_scan_all, self.page_list)
        self.setTabOrder(self.page_list, btn_delete_page)
        self.setTabOrder(btn_delete_page, btn_clear_pages)
        self.setTabOrder(btn_clear_pages, self.btn_ocr)
        self.setTabOrder(self.btn_ocr, self.btn_save_document)
        self.setTabOrder(self.btn_save_document, self.btn_read)
        self.setTabOrder(self.btn_read, btn_export_img)
        self.setTabOrder(btn_export_img, self.btn_import_pdf)

        # Uložit reference pro testování Tab order a pro focus handling
        self._scroll_area = scroll
        self._scroll_container = container
        # QGroupBox nesmí být focusable – pojistka (default NoFocus, ale explicitně)
        for _gb in (ocr_group, scan_group, macro_group):
            _gb.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        # Dekorativní QLabel také ne
        for _lbl in (lbl_lang, lbl_engine, lbl_diac, lbl_device, lbl_source, lbl_dpi, lbl_color):
            _lbl.setFocusPolicy(Qt.FocusPolicy.NoFocus)

    # ---------- Počáteční fokus pro NVDA ----------
    def showEvent(self, event) -> None:
        super().showEvent(event)
        if not self._initial_focus_done:
            # Odložit na konec event loopu – okno musí být viditelné a mapped.
            # QTimer.singleShot(0, ...) je robustnější než přímé setFocus v __init__.
            QTimer.singleShot(0, self._set_initial_focus)

    def _set_initial_focus(self) -> None:
        if self._initial_focus_done:
            return
        # Pokud okno ještě není aktivní (Windows activation delay), zkusit znovu
        # – neagresivně, bez krádeže fokusu. Max 5 pokusů po 100 ms.
        if not self.isActiveWindow():
            # Zjistit, zda už byl pokus; uložit counter do atributu
            tries = getattr(self, "_initial_focus_tries", 0)
            if tries < 5:
                self._initial_focus_tries = tries + 1
                QTimer.singleShot(100, self._set_initial_focus)
                return
            # Po vyčerpání pokusů nastavit fokus i bez isActiveWindow –
            # QWidget.setFocus() vyžaduje active window pro NVDA, ale
            # fokus bude doručen jakmile se okno aktivuje uživatelem/WM.
            # Nepoužívat QApplication.setActiveWindow() (deprecated) ani
            # agresivní activateWindow() pokud uživatel mezitím přešel jinam.
        # Nepoužívat zastaralé QApplication.setActiveWindow().
        # activateWindow() pouze pokud je to bezpečné a okno je viditelné,
        # ale nekrást fokus – proto jen pokud je okno viditelné a ještě neaktivní
        # a uživatel zjevně aplikaci právě spustil (první show).
        # V praxi stačí setFocus; Windows dá aktivaci automaticky při spuštění.
        self.lang_combo.setFocus(Qt.FocusReason.OtherFocusReason)
        # Qt interně pošle QAccessible::Focus event přes UIA bridge.
        # Explicitní QAccessible.updateAccessibility() by bylo duplicitní
        # a v PyQt6 6.11 není QAccessible v Python API vůbec exponován.
        # Pokus o podmíněné poslání pouze pokud by bylo dostupné:
        try:
            import importlib
            qacc = importlib.import_module("PyQt6.QtGui")
            QAccessible = getattr(qacc, "QAccessible", None)
            if QAccessible is not None and hasattr(QAccessible, "updateAccessibility"):
                # isActive() guard – posílat jen když běží AT
                is_active = getattr(QAccessible, "isActive", None)
                if is_active is None or is_active():
                    QAccessible.updateAccessibility(
                        self.lang_combo, 0, QAccessible.Event.Focus
                    )
        except Exception:
            pass
        self._initial_focus_done = True

    def _setup_shortcuts(self) -> None:
        QShortcut(QKeySequence("Ctrl+S"), self).activated.connect(self.scan_pages)
        QShortcut(QKeySequence("Ctrl+Shift+N"), self).activated.connect(self.scan_next_page)
        QShortcut(QKeySequence("Ctrl+Shift+S"), self).activated.connect(self.scan_all_pages)
        QShortcut(QKeySequence("Ctrl+O"), self).activated.connect(self.run_ocr)
        QShortcut(QKeySequence("Ctrl+U"), self).activated.connect(self.save_document)
        QShortcut(QKeySequence("Ctrl+P"), self).activated.connect(self.read_last_text)
        QShortcut(QKeySequence("Ctrl+E"), self).activated.connect(self.export_images)
        QShortcut(QKeySequence("Ctrl+I"), self).activated.connect(self.import_pdfs)
        QShortcut(QKeySequence("Ctrl+Q"), self).activated.connect(self.close)
        QShortcut(QKeySequence("Delete"), self).activated.connect(self.delete_page)
        QShortcut(QKeySequence("Ctrl+M"), self).activated.connect(self._open_macro_editor)
        self._setup_macro_shortcuts()

    def _setup_macro_shortcuts(self) -> None:
        for sc in self._macro_shortcuts:
            sc.setEnabled(False)
            sc.deleteLater()
        self._macro_shortcuts = []
        for macro in self.macro_manager.macros:
            if macro.shortcut:
                sc = QShortcut(QKeySequence(macro.shortcut), self)
                sc.activated.connect(lambda checked=False, m=macro: self._run_macro(m))
                self._macro_shortcuts.append(sc)

    # ---------- Settings persistence ----------
    def _load_settings(self) -> None:
        lang_idx = self.settings.value("ocr/lang_index", 0, type=int)
        if 0 <= lang_idx < self.lang_combo.count():
            self.lang_combo.setCurrentIndex(lang_idx)
        engine_idx = self.settings.value("ocr/engine_index", 0, type=int)
        if 0 <= engine_idx < self.engine_combo.count():
            self.engine_combo.setCurrentIndex(engine_idx)
        dpi_val = self.settings.value("scan/dpi", 300, type=int)
        try:
            dpi_val = int(dpi_val)
        except (TypeError, ValueError):
            dpi_val = 300
        _allowed_dpi = {200, 300, 400, 600, 1200}
        if dpi_val not in _allowed_dpi:
            dpi_val = 300
        idx = self.dpi_combo.findData(dpi_val)
        if idx >= 0:
            self.dpi_combo.setCurrentIndex(idx)
        else:
            idx_def = self.dpi_combo.findData(300)
            if idx_def >= 0:
                self.dpi_combo.setCurrentIndex(idx_def)
        source_idx = self.settings.value("scan/source_index", 0, type=int)
        if 0 <= source_idx < self.source_combo.count():
            self.source_combo.setCurrentIndex(source_idx)
        color_idx = self.settings.value("scan/color_index", 0, type=int)
        if 0 <= color_idx < self.color_combo.count():
            self.color_combo.setCurrentIndex(color_idx)
        self.batch_cb.setChecked(self.settings.value("scan/batch", False, type=bool))
        self.preprocess_cb.setChecked(self.settings.value("ocr/preprocess", True, type=bool))
        self.diacritics_cb.setChecked(self.settings.value("ocr/diacritics", False, type=bool))

    # ---------- EasyOCR preload (startup, změna jazyka, retry) ----------
    def _easyocr_lang_for_current_settings(self) -> str:
        raw = self.lang_combo.currentText() if hasattr(self, "lang_combo") else "ces"
        pyt_code = raw.split()[0].strip()
        return LANG_MAP_EASYOCR.get(pyt_code, pyt_code[:2])

    def _preload_easyocr_if_needed(self, force: bool = False) -> None:
        """Připraví EasyOCR na pozadí pokud je zvolen jako engine.
        Nespouští se v GUI vlákně, používá EasyOCRPreloadThread.
        """
        try:
            engine = self.engine_combo.currentText() if hasattr(self, "engine_combo") else ""
        except Exception:
            engine = ""
        # Preferované chování: připravuj pouze pokud je EasyOCR aktuálně zvolený
        if engine != "EasyOCR" and not force:
            # Informovat pouze že není potřeba, ale nehlásit chybu
            self._update_easyocr_status_label("idle", "EasyOCR není zvolen.")
            logger.info("EasyOCR preload skip – engine=%s", engine)
            return
        easy_lang = self._easyocr_lang_for_current_settings()
        # Pokud už ready → pouze informuj, nestahuj znovu
        if is_easyocr_ready(easy_lang):
            self._update_easyocr_status_label("ready", f"EasyOCR je připraven ({easy_lang}).")
            # Potlačit duplicitní hlas pokud už bylo oznámeno
            logger.info("EasyOCR already ready lang=%s", easy_lang)
            return
        if is_easyocr_preparing(easy_lang):
            self._update_easyocr_status_label("preparing", "EasyOCR se připravuje.")
            logger.info("EasyOCR already preparing lang=%s", easy_lang)
            return
        # Pokud předchozí thread ještě běží (jiný jazyk), počkej na dokončení – nezahazuj
        if self._easyocr_preload_thread is not None and self._easyocr_preload_thread.isRunning():
            logger.info("EasyOCR preload already running, queued lang=%s", easy_lang)
            # Necháme doběhnout předchozí, nový jazyk se načte při příštím OCR nebo po dokončení
            # Ale pro změnu jazyka spustíme nový pokud je jiný jazyk
            # Aby nedošlo k souběhu stejného modelu, zkontroluj lang
            # Pokud je to jiný jazyk než aktuální thread, spustíme nový po dokončení
            # Pro jednoduchost: nezahajuj druhý pokud první běží
            self._update_easyocr_status_label("preparing", "EasyOCR se připravuje.")
            return
        # Zahájit přípravu na pozadí
        logger.info("EasyOCR preload start lang=%s engine=%s", easy_lang, engine)
        self._update_easyocr_status_label("preparing", "Připravuji OCR.")
        speak("Připravuji OCR.")
        # Informovat zda bude stahování (pouze pro log/hlas)
        try:
            if not _is_easyocr_model_cached(easy_lang):
                # hlas pro stahování přijde z worker signálu, ale můžeme připravit
                logger.info("EasyOCR model not cached, will download lang=%s", easy_lang)
        except Exception:
            pass
        # Vytvoř nový worker
        thread = EasyOCRPreloadThread(easy_lang)
        thread.started.connect(self._on_easyocr_started)
        thread.downloading.connect(self._on_easyocr_downloading)
        thread.finished_ok.connect(self._on_easyocr_ready)
        thread.failed.connect(self._on_easyocr_failed)
        thread.progress_msg.connect(self._on_easyocr_progress_msg)
        # Uložení reference aby nebyl GC
        self._easyocr_preload_thread = thread
        thread.finished.connect(lambda: self._on_easyocr_thread_finished(thread))
        thread.start()

    def _on_easyocr_thread_finished(self, thread) -> None:
        # Uklidit referenci pokud je to aktuální thread
        try:
            if self._easyocr_preload_thread is thread:
                # Ponech referenci krátce, pak vyčisti – ale nech signály doběhnout
                pass
        except Exception:
            pass

    def _update_easyocr_status_label(self, state: str, text: str) -> None:
        self._easyocr_status_text = text
        try:
            if hasattr(self, "easyocr_status_label") and self.easyocr_status_label is not None:
                # Map state na přístupný text
                if state == "ready":
                    label = "EasyOCR je připraven."
                elif state == "preparing":
                    label = "EasyOCR se připravuje."
                elif state == "failed":
                    label = "EasyOCR se nepodařilo připravit."
                elif state == "idle":
                    label = "EasyOCR není zvolen."
                else:
                    label = text
                self.easyocr_status_label.setText(label)
                self.easyocr_status_label.setAccessibleName("Stav EasyOCR")
                self.easyocr_status_label.setAccessibleDescription(label)
                # Pro NVDA – změna textu je oznámena; doplň speak kde je vhodné
        except Exception:
            pass
        # řízení viditelnosti retry tlačítka
        try:
            if hasattr(self, "btn_retry_easyocr"):
                if state == "failed":
                    self.btn_retry_easyocr.show()
                    self.btn_retry_easyocr.setEnabled(True)
                elif state == "ready":
                    self.btn_retry_easyocr.hide()
                elif state == "preparing":
                    self.btn_retry_easyocr.hide()
                elif state == "idle":
                    self.btn_retry_easyocr.hide()
        except Exception:
            pass

    def _on_easyocr_started(self, lang: str) -> None:
        logger.info("EasyOCR started lang=%s", lang)
        self._update_easyocr_status_label("preparing", "EasyOCR se připravuje.")
        speak("Připravuji OCR.")

    def _on_easyocr_downloading(self, lang: str) -> None:
        logger.info("EasyOCR downloading lang=%s", lang)
        self._update_easyocr_status_label("preparing", "Stahuji model pro EasyOCR.")
        speak("Stahuji model pro EasyOCR.")

    def _on_easyocr_ready(self, lang: str) -> None:
        logger.info("EasyOCR ready lang=%s", lang)
        self._update_easyocr_status_label("ready", "EasyOCR je připraven.")
        speak("EasyOCR je připraven.")
        # Pokud byl OCR požadavek odložen během přípravy, spusť jej nyní
        if self._easyocr_pending_ocr is not None:
            pending = self._easyocr_pending_ocr
            self._easyocr_pending_ocr = None
            logger.info("EasyOCR running pending OCR after ready lang=%s", lang)
            # pending obsahuje {"mode": "full"/"incremental", "start_idx": int}
            try:
                if pending.get("mode") == "incremental":
                    self._run_ocr_incremental(pending.get("start_idx", 0))
                else:
                    self._run_ocr_flow(interactive=pending.get("interactive", True))
            except Exception as e:
                logger.exception("Pending OCR failed after preload: %s", e)

    def _on_easyocr_failed(self, lang: str, err: str) -> None:
        logger.error("EasyOCR failed lang=%s err=%s", lang, err)
        self._update_easyocr_status_label("failed", "EasyOCR se nepodařilo připravit.")
        speak("EasyOCR se nepodařilo připravit. Zkontrolujte připojení k internetu a zkuste to znovu.")
        # pending OCR zůstane – uživatel musí explicitně opakovat nebo kliknout retry
        # ale nezahazuj pending, aby mohl po retry pokračovat
        # QMessageBox nezobrazovat automaticky při startu – pouze hlas a label

    def _on_easyocr_progress_msg(self, msg: str) -> None:
        # Nepřehlušovat – pouze log
        logger.info("EasyOCR progress_msg: %s", msg)

    def _on_retry_easyocr_clicked(self) -> None:
        logger.info("EasyOCR retry clicked")
        speak("Připravuji EasyOCR znovu.")
        self._preload_easyocr_if_needed(force=True)
        # Po kliknutí přesuň fokus zpět na stavový label pro potvrzení
        try:
            self.easyocr_status_label.setFocus()
        except Exception:
            pass

    def _on_engine_or_lang_changed(self) -> None:
        # Uložit nastavení a případně spustit preload pro nový jazyk
        try:
            self._save_settings()
        except Exception:
            pass
        # Pokud je zvolen EasyOCR, připrav nový jazyk na pozadí (pokud není cached)
        try:
            engine = self.engine_combo.currentText()
            if engine == "EasyOCR":
                # Nezahlcovat – preload zkontroluje ready/preparing
                QTimer.singleShot(100, self._preload_easyocr_if_needed)
            else:
                self._update_easyocr_status_label("idle", "EasyOCR není zvolen.")
        except Exception:
            pass

    def _on_diacritics_toggled(self, checked: bool) -> None:
        # Hlasová odezva při přepnutí
        if checked:
            speak("Oprava české diakritiky zapnuta.")
        else:
            speak("Oprava české diakritiky vypnuta.")
        self._save_settings()

    def _save_settings(self) -> None:
        self.settings.setValue("ocr/lang_index", self.lang_combo.currentIndex())
        self.settings.setValue("ocr/engine_index", self.engine_combo.currentIndex())
        self.settings.setValue("scan/dpi", self.dpi_combo.currentData())
        self.settings.setValue("scan/source_index", self.source_combo.currentIndex())
        self.settings.setValue("scan/color_index", self.color_combo.currentIndex())
        self.settings.setValue("scan/batch", self.batch_cb.isChecked())
        self.settings.setValue("ocr/preprocess", self.preprocess_cb.isChecked())
        self.settings.setValue("ocr/diacritics", self.diacritics_cb.isChecked())

    def closeEvent(self, event) -> None:
        self._save_settings()
        super().closeEvent(event)

    # ---------- Device management ----------
    def reload_devices(self) -> None:
        try:
            devices = self.scanner.list_devices()
            self.device_combo.clear()
            for d in devices:
                self.device_combo.addItem(d)
            if devices:
                self.scanner.connect_device(devices[0])
            elif not self.scanner.naps2_exe:
                QMessageBox.warning(
                    self, "NAPS2 nenalezen",
                    "NAPS2.Console.exe nebyl nalezen. Ujistěte se, že je NAPS2 nainstalován."
                )
        except Exception as e:
            QMessageBox.critical(self, "Chyba", f"Nenašel jsem žádný skener:\n{e}")

    # ---------- Scanning ----------
    def scan_pages(self) -> None:
        # If pages already exist, never clear without confirmation.
        if self.scanned_images:
            count = len(self.scanned_images)
            speak(f"Máte již {count} naskenovaných stránek. Potvrďte smazání nebo přidejte další stránku.")
            msg = QMessageBox(self)
            msg.setWindowTitle("Nový dokument")
            msg.setText(f"Máte již {count} naskenovaných stránek. Chcete začít nový dokument?")
            msg.setInformativeText("Stávající stránky budou smazány. Můžete také přidat další stránku ke stávajícímu dokumentu.")
            btn_new = msg.addButton("Ano, smazat a začít znovu", QMessageBox.ButtonRole.YesRole)
            btn_new.setAccessibleName("Ano, smazat stávající stránky a začít nový dokument")
            btn_append = msg.addButton("Přidat další stránku", QMessageBox.ButtonRole.ActionRole)
            btn_append.setAccessibleName("Přidat další stránku ke stávajícímu dokumentu")
            btn_cancel = msg.addButton("Zrušit", QMessageBox.ButtonRole.NoRole)
            btn_cancel.setAccessibleName("Zrušit, zachovat stránky")
            msg.setDefaultButton(btn_cancel)
            msg.exec()
            clicked = msg.clickedButton()
            if clicked == btn_append:
                self.scan_next_page()
                return
            if clicked != btn_new:
                speak("Skenování zrušeno, stránky zachovány.")
                self.page_list.setFocus()
                return
            # User confirmed new document – clear everything
            self.scanned_images.clear()
            self.raw_scanned_images.clear()
            self.page_list.clear()
            self.last_ocr_results.clear()
            self._last_text = ""
            self._last_original_text = ""
            self._last_edited_pages = None
            self._diacritics_failed = False
            self.btn_ocr.setEnabled(False)
            self.btn_read.setEnabled(False)
            self._update_save_button_state()
        else:
            # No pages – ensure clean state (in case of residual OCR results)
            self.last_ocr_results.clear()
            self._last_text = ""
            self._last_original_text = ""
            self._last_edited_pages = None
            self._diacritics_failed = False

        speak("Zahajuji skenování.")
        # Keep button states disabled until at least one page succeeds
        self.btn_ocr.setEnabled(False)
        self.btn_read.setEnabled(False)
        self._update_save_button_state()

        # Remember count before scanning for correct voice after
        start_count = len(self.scanned_images)

        while True:
            self.last_scanned_image = None
            self._scan_cancelled = False
            self._do_scan_single()

            if self._scan_cancelled or self.last_scanned_image is None:
                break

            if self.batch_cb.isChecked():
                continue

            speak("Další stránka?")
            msg = QMessageBox(self)
            msg.setWindowTitle("Další stránka")
            msg.setText("Další stránka?")
            btn_yes = msg.addButton("Ano", QMessageBox.ButtonRole.YesRole)
            btn_yes.setAccessibleName("Ano, skenovat další stránku")
            btn_no = msg.addButton("Ne", QMessageBox.ButtonRole.NoRole)
            btn_no.setAccessibleName("Ne, ukončit skenování")
            msg.exec()
            if msg.clickedButton() != btn_yes:
                break

        if self.scanned_images:
            self.btn_ocr.setEnabled(True)
            self._update_save_button_state()
            # In new-document mode we cleared page_list, so repopulate.
            # If start_count was 0 we are in new mode; otherwise we already cleared.
            self.page_list.clear()
            self.page_list.addItems([f"Stránka {i+1}" for i in range(len(self.scanned_images))])
            self.page_list.setFocus()
            speak(f"Skenování dokončeno. {len(self.scanned_images)} stránek.")
        else:
            self.btn_scan.setFocus()
            self._update_save_button_state()
            speak("Skenování dokončeno, žádné stránky.")

    def scan_next_page(self) -> None:
        speak("Přidávám další stránku ke stávajícímu dokumentu.")
        count_before = len(self.scanned_images)
        self.last_scanned_image = None
        self._scan_cancelled = False
        self._do_scan_single()

        if self._scan_cancelled or self.last_scanned_image is None:
            speak("Skenování další stránky zrušeno.")
            if self.scanned_images:
                self.page_list.setFocus()
            else:
                self.btn_scan.setFocus()
            return

        # _scan_finished already appended images; now update page_list incrementally
        new_num = len(self.scanned_images)
        # Guard against double-add if scan_pages bulk path was used – but scan_next_page
        # never clears, so page_list should have count_before items
        if self.page_list.count() < new_num:
            self.page_list.addItem(f"Stránka {new_num}")
        else:
            # Fallback: rebuild to stay consistent
            self.page_list.clear()
            self.page_list.addItems([f"Stránka {i+1}" for i in range(new_num)])

        self.btn_ocr.setEnabled(True)
        self._update_save_button_state()
        self.page_list.setCurrentRow(new_num - 1)
        self.page_list.setFocus()
        total = len(self.scanned_images)
        # Keep existing OCR results for pages 1..count_before; new page is not OCRed yet
        if self.last_ocr_results and len(self.last_ocr_results) < total:
            speak(f"Stránka {new_num} přidána. Celkem {total} stránek. OCR nové stránky zatím nebylo provedeno.")
        else:
            speak(f"Stránka {new_num} přidána. Celkem {total} stránek.")

    def _renumber_pages(self) -> None:
        for i in range(self.page_list.count()):
            self.page_list.item(i).setText(f"Stránka {i+1}")

    # ---------- Helpers pro stav OCR a tlačítko Uložit ----------
    def _has_ocr_text(self) -> bool:
        for page in self.last_ocr_results:
            for item in page:
                if item.text and item.text.strip():
                    return True
        # fallback – _last_text může obsahovat text i bez bbox
        if self._last_text and self._last_text.strip():
            # ověř že text není jen headery
            stripped = re.sub(r"---\s*Stránka\s+\d+\s*---", "", self._last_text).strip()
            return bool(stripped)
        return False

    def _is_ocr_empty(self) -> bool:
        return not self.last_ocr_results or not self._has_ocr_text()

    def _is_ocr_complete(self) -> bool:
        total = len(self.scanned_images)
        if total == 0:
            return False
        if len(self.last_ocr_results) != total:
            return False
        return self._has_ocr_text()

    def _missing_ocr_count(self) -> int:
        return max(0, len(self.scanned_images) - len(self.last_ocr_results))

    def _update_save_button_state(self) -> None:
        has_pages = len(self.scanned_images) > 0
        if hasattr(self, "btn_save_document"):
            self.btn_save_document.setEnabled(has_pages)

    def scan_all_pages(self) -> None:
        old_batch = self.batch_cb.isChecked()
        self.batch_cb.setChecked(True)
        self.scan_pages()
        self.batch_cb.setChecked(old_batch)

    def _do_scan_single(self) -> None:
        device_name = self.device_combo.currentText()
        if not device_name:
            QMessageBox.warning(self, "Chyba", "Není vybrán žádný skener.")
            return

        self.scanner.connect_device(device_name)
        dpi = self.dpi_combo.currentData()
        if dpi is None:
            dpi = 300
        source = self.source_combo.currentText()
        color_text = self.color_combo.currentText()
        color_mode = (
            "Color" if color_text == "Barevný"
            else "Grayscale" if color_text == "Šedý"
            else "BlackWhite"
        )

        self.progress_dialog = QProgressDialog(
            "Skenuji (přes NAPS2)...", "Zrušit", 0, 0, self
        )
        self.progress_dialog.canceled.connect(self._cancel_scan)
        self.progress_dialog.show()

        self.scan_thread = ScanThread(self.scanner, dpi, color_mode, source)
        self.scan_thread.finished.connect(self._scan_finished)
        self.scan_thread.error.connect(self._scan_error)
        self.scan_thread.start()

        loop = QEventLoop()
        self.scan_thread.finished.connect(loop.quit)
        self.scan_thread.error.connect(loop.quit)
        loop.exec()

        self.progress_dialog.cancel()

    def _cancel_scan(self) -> None:
        self._scan_cancelled = True
        speak("Skenování zrušeno.")

    def _scan_finished(self, img: Image.Image) -> None:
        self.progress_dialog.cancel()
        if self._scan_cancelled:
            self.last_scanned_image = None
            return

        preview = PreviewDialog(img, len(self.scanned_images) + 1)
        if preview.exec():
            original = preview.image
            processed = original
            if self.preprocess_cb.isChecked():
                processed = preprocess_image(original)
            self.raw_scanned_images.append(original)
            self.scanned_images.append(processed)
            self.last_scanned_image = processed
        else:
            self.last_scanned_image = None

    def _scan_error(self, msg: str) -> None:
        self.progress_dialog.cancel()
        speak(f"Chyba skenování: {msg}")
        QMessageBox.critical(self, "Chyba skenování", msg)
        self.last_scanned_image = None

    # ---------- Page management ----------
    def delete_page(self) -> None:
        row = self.page_list.currentRow()
        if row < 0:
            speak("Není vybrána žádná stránka k smazání.")
            return
        page_num = row + 1
        speak(f"Potvrdit smazání stránky {page_num}.")
        msg = QMessageBox(self)
        msg.setWindowTitle("Potvrdit smazání stránky")
        msg.setText(f"Opravdu chcete smazat stránku {page_num}?")
        msg.setInformativeText("Tuto akci nelze vrátit.")
        btn_yes = msg.addButton("Ano", QMessageBox.ButtonRole.YesRole)
        btn_yes.setAccessibleName(f"Ano, smazat stránku {page_num}")
        btn_no = msg.addButton("Ne", QMessageBox.ButtonRole.NoRole)
        btn_no.setAccessibleName("Ne, ponechat stránku")
        msg.setDefaultButton(btn_no)
        msg.exec()
        if msg.clickedButton() != btn_yes:
            speak("Smazání zrušeno.")
            self.page_list.setFocus()
            self.page_list.setCurrentRow(row)
            return

        self.page_list.takeItem(row)
        del self.scanned_images[row]
        del self.raw_scanned_images[row]
        if self.last_ocr_results and row < len(self.last_ocr_results):
            del self.last_ocr_results[row]
        if self._last_edited_pages is not None and row < len(self._last_edited_pages):
            del self._last_edited_pages[row]
        # Pokud byl _last_text s headery, přegeneruj ho podle zbývajících stránek
        if self._last_edited_pages is not None:
            # Přegeneruj _last_text aby reflektoval smazání stránky
            parts = []
            for idx, p in enumerate(self._last_edited_pages):
                parts.append(f"--- Stránka {idx+1} ---\n{p}\n\n")
            self._last_text = "".join(parts)
        # Renumber remaining items to keep Stránka 1..N consistent
        self._renumber_pages()
        remaining = len(self.scanned_images)
        if remaining == 0:
            self.btn_ocr.setEnabled(False)
            self.btn_read.setEnabled(False)
            self._last_text = ""
            self._last_original_text = ""
            self._last_edited_pages = None
            self._diacritics_failed = False
            self._update_save_button_state()
            speak("Všechny stránky smazány.")
            self.btn_scan.setFocus()
        else:
            speak(f"Stránka smazána. Zbývá {remaining} stránek.")
            self._update_save_button_state()
            # Focus na stejnou pozici nebo poslední
            next_row = min(row, remaining - 1)
            self.page_list.setCurrentRow(next_row)
            self.page_list.setFocus()

    def clear_pages(self) -> None:
        count = len(self.scanned_images)
        if count == 0:
            speak("Žádné stránky ke smazání.")
            return
        speak(f"Potvrdit smazání všech {count} stránek.")
        msg = QMessageBox(self)
        msg.setWindowTitle("Potvrdit smazání všech stránek")
        msg.setText(f"Opravdu chcete smazat všech {count} naskenovaných stránek?")
        msg.setInformativeText("Budou odstraněny všechny naskenované stránky a jejich OCR výsledky. Tuto akci nelze vrátit.")
        btn_yes = msg.addButton("Ano", QMessageBox.ButtonRole.YesRole)
        btn_yes.setAccessibleName("Ano, smazat všechny stránky")
        btn_no = msg.addButton("Ne", QMessageBox.ButtonRole.NoRole)
        btn_no.setAccessibleName("Ne, ponechat stránky")
        msg.setDefaultButton(btn_no)
        msg.exec()
        if msg.clickedButton() != btn_yes:
            speak("Smazání zrušeno.")
            self.page_list.setFocus()
            return
        self.page_list.clear()
        self.scanned_images.clear()
        self.raw_scanned_images.clear()
        self.last_ocr_results.clear()
        self._last_text = ""
        self._last_original_text = ""
        self._last_edited_pages = None
        self._diacritics_failed = False
        self.btn_ocr.setEnabled(False)
        self.btn_read.setEnabled(False)
        self._update_save_button_state()
        speak(f"Všech {count} stránek smazáno.")
        self.btn_scan.setFocus()

    # ---------- OCR – guard pro EasyOCR preload ----------
    def _handle_easyocr_guard_before_ocr(self, engine: str, lang_raw: str, pending_info: dict) -> bool:
        """Vrátí True pokud lze pokračovat, False pokud byl OCR odložen nebo zablokován.
        Zajišťuje: žádný duplicitní Reader, žádný pád, přístupná hláška.
        """
        if engine != "EasyOCR":
            return True
        easy_lang = LANG_MAP_EASYOCR.get(lang_raw.split()[0].strip(), lang_raw.split()[0].strip()[:2])
        if is_easyocr_preparing(easy_lang):
            # Uložit pending pouze pokud ještě není
            if self._easyocr_pending_ocr is None:
                self._easyocr_pending_ocr = pending_info
            speak("EasyOCR se ještě připravuje. OCR bude spuštěno po dokončení přípravy.")
            self._update_easyocr_status_label("preparing", "EasyOCR se ještě připravuje. OCR bude spuštěno po dokončení přípravy.")
            # Zajistit že preload běží (pokud z nějakého důvodu neběží)
            if self._easyocr_preload_thread is None or not self._easyocr_preload_thread.isRunning():
                # Pro jistotu znovu spustit
                try:
                    self._preload_easyocr_if_needed()
                except Exception:
                    pass
            return False
        if is_easyocr_failed(easy_lang):
            speak("EasyOCR se nepodařilo připravit. Zkontrolujte připojení k internetu a zkuste to znovu.")
            # Ukaž retry tlačítko
            self._update_easyocr_status_label("failed", "EasyOCR se nepodařilo připravit.")
            QMessageBox.warning(self, "EasyOCR není připraven", "EasyOCR se nepodařilo připravit. Zkontrolujte připojení k internetu a zkuste to znovu.\n\nKlikněte na 'Připravit EasyOCR znovu'.")
            return False
        if not is_easyocr_ready(easy_lang):
            # Není připraven a není preparing ani failed -> zahájit přípravu a odložit OCR
            if self._easyocr_pending_ocr is None:
                self._easyocr_pending_ocr = pending_info
            speak("EasyOCR se ještě připravuje. OCR bude spuštěno po dokončení přípravy.")
            self._update_easyocr_status_label("preparing", "EasyOCR se připravuje.")
            try:
                self._preload_easyocr_if_needed(force=True)
            except Exception:
                pass
            return False
        return True

    # ---------- OCR ----------
    def _run_ocr_incremental(self, start_idx: int) -> None:
        """OCR pouze pro nově přidané stránky start_idx..end, zachová 1..start_idx-1."""
        new_images = self.scanned_images[start_idx:]
        if not new_images:
            return
        # Guard – pokud EasyOCR není ready, odlož
        _eng = self.engine_combo.currentText() if hasattr(self, "engine_combo") else ""
        _lang_raw = self.lang_combo.currentText() if hasattr(self, "lang_combo") else "ces"
        if not self._handle_easyocr_guard_before_ocr(_eng, _lang_raw, {"mode": "incremental", "start_idx": start_idx}):
            return
        total_new = len(new_images)
        total_all = len(self.scanned_images)
        self._ocr_total_pages = total_all
        self._ocr_last_announced_page = start_idx  # aby progress hlásil od start_idx+1
        speak(f"Zahajuji OCR {total_new} nových stránek (celkem {total_all}).")
        base_engine = self.engine_combo.currentText()
        base_lang = self.lang_combo.currentText()
        # Jednorázový pokus pro nové stránky – stačí base pokus, fallback ponechán pro full flow
        # Zde použijeme stejný mechanismus jako _run_ocr_once ale s merge
        old_results = list(self.last_ocr_results)
        old_text = self._last_text

        # Dočasné sloty pro merge
        new_results_holder: list[list[OcrResult]] = []
        new_text_holder: list[str] = []

        loop = QEventLoop()

        def on_results(r):
            new_results_holder.clear()
            new_results_holder.extend(r)

        def on_finished(txt: str):
            new_text_holder.append(txt)
            loop.quit()

        # Vytvoř thread pro nové stránky
        diac_enabled = self.diacritics_cb.isChecked() if hasattr(self, "diacritics_cb") else False
        self._diacritics_failed = False
        self._diacritics_enabled_cache = diac_enabled
        self.ocr_thread = create_ocr_thread(base_engine, new_images, base_lang, diacritics_enabled=diac_enabled)
        self.ocr_thread.ocr_results.connect(on_results)
        self.ocr_thread.finished.connect(on_finished)
        if hasattr(self.ocr_thread, "diacritics_failed"):
            self.ocr_thread.diacritics_failed.connect(self._on_diacritics_failed)
        if hasattr(self.ocr_thread, "failed"):
            self.ocr_thread.failed.connect(self._on_easyocr_thread_failed)
        if base_engine == "EasyOCR":
            self.progress_dialog = QProgressDialog("Načítám EasyOCR model...", "Zrušit", 0, 0, self)
            if hasattr(self.ocr_thread, "model_loading"):
                self.ocr_thread.model_loading.connect(self._on_model_loaded)
        else:
            self.progress_dialog = QProgressDialog("Probíhá OCR (Tesseract)...", "Zrušit", 0, 100, self)
        self.ocr_thread.progress.connect(self._on_ocr_progress)
        self.progress_dialog.show()
        self.ocr_thread.start()
        loop.exec()
        self.progress_dialog.cancel()

        if not new_results_holder:
            # Pokud EasyOCR selhal kvůli chybě modelu, nehlásit jen prázdný výsledek
            if base_engine == "EasyOCR":
                easy_lang = LANG_MAP_EASYOCR.get(base_lang.split()[0].strip(), base_lang.split()[0].strip()[:2])
                if is_easyocr_failed(easy_lang):
                    speak("EasyOCR se nepodařilo připravit. Zkontrolujte připojení k internetu a zkuste to znovu.")
                    return
            speak("OCR nových stránek neprodukovalo žádné výsledky.")
            return
        # Oprav číslování stránek v novém textu (thread čísluje od 1)
        # Sestav korektní full_text pro nové stránky s posunutým číslováním
        corrected_new_text = ""
        for i, page_results in enumerate(new_results_holder):
            page_num = start_idx + i + 1
            # Extrahuj text stránky z new_text_holder[0] – jednodušší rekonstruovat z page_results
            # ale zachovej původní diakritizovaný text včetně odřádkování – použij display_text
            page_text_parts = [r.display_text for r in page_results]
            # Pokud page_results obsahuje věty s newline, spoj mezery
            page_text = "\n".join(page_text_parts)
            # Pokud původní new_text_holder obsahuje více, použij ho přímo s přečíslováním
            # Pro jednoduchost použij page_text z results
            corrected_new_text += f"--- Stránka {page_num} ---\n{page_text}\n\n"

        # Merge
        self.last_ocr_results = old_results + new_results_holder
        # Pokud byl původní full_text s headery, připoj nové; jinak použij corrected
        # Pro zachování přesnosti použij starý text + corrected_new_text
        if old_text and old_text.strip():
            self._last_text = old_text.rstrip() + "\n\n" + corrected_new_text
        else:
            self._last_text = corrected_new_text
        # Aktualizuj per-page
        try:
            pages = _split_text_by_page_headers(self._last_text)
            if pages is not None and len(pages) == len(self.last_ocr_results):
                self._last_edited_pages = pages
            else:
                self._last_edited_pages = None
        except Exception:
            self._last_edited_pages = None
        # Rekonstrukce originálu
        try:
            orig_parts = []
            for idx, page in enumerate(self.last_ocr_results):
                if page:
                    texts = [r.original_text if r.original_text is not None else r.text for r in page]
                    orig_parts.append(f"--- Stránka {idx + 1} ---\n" + "\n".join(texts) + "\n\n")
            self._last_original_text = "".join(orig_parts) if any(orig_parts) else self._last_text
        except Exception:
            self._last_original_text = self._last_text
        self.btn_read.setEnabled(True)
        # Hlasová odezva
        from diacritics import should_diacritize
        lang = self.lang_combo.currentText() if hasattr(self, "lang_combo") else ""
        diac_active = should_diacritize(lang, self._diacritics_enabled_cache)
        if self._diacritics_failed and diac_active:
            speak("OCR dokončeno. Oprava české diakritiky se nepodařila. Byl zachován původní text.")
        elif diac_active:
            speak("OCR dokončeno a česká diakritika opravena.")
        else:
            speak(f"OCR dokončeno pro {total_all} stránek.")

        self._ocr_mode = "interactive"
        # Pokud je pending save (voláno z Uložit dokument), náhled nezobrazuj – rovnou pokračuj k uložení
        if getattr(self, "_pending_save_format", None):
            return
        self._show_ocr_preview(self._last_text, is_incremental=True, incremental_start_idx=start_idx)

    def run_ocr(self) -> None:
        if not self.scanned_images:
            QMessageBox.warning(self, "Chyba", "Nejdříve naskenujte stránky.")
            return
        total = len(self.scanned_images)
        # Inkrementální OCR: pokud již máme OCR pro část stránek, zpracuj pouze nové
        if self.last_ocr_results and 0 < len(self.last_ocr_results) < total:
            # Guard se provede uvnitř _run_ocr_incremental
            self._run_ocr_incremental(len(self.last_ocr_results))
            return
        self._ocr_total_pages = total
        self._ocr_last_announced_page = 0
        # Guard pro EasyOCR před zahájením flow
        _eng = self.engine_combo.currentText() if hasattr(self, "engine_combo") else ""
        _lang_raw = self.lang_combo.currentText() if hasattr(self, "lang_combo") else "ces"
        if not self._handle_easyocr_guard_before_ocr(_eng, _lang_raw, {"mode": "full", "interactive": True}):
            return
        speak(f"Zahajuji OCR {total} stránek.")
        self._run_ocr_flow(interactive=True)

    def _on_easyocr_thread_failed(self, msg: str) -> None:
        logger.error("EasyOCR thread failed: %s", msg)
        # progress_dialog canceled handled in _ocr_finished / loop
        speak("EasyOCR se nepodařilo připravit. Zkontrolujte připojení k internetu a zkuste to znovu.")
        self._update_easyocr_status_label("failed", "EasyOCR se nepodařilo připravit.")
        try:
            if hasattr(self, "progress_dialog") and self.progress_dialog is not None:
                self.progress_dialog.cancel()
        except Exception:
            pass

    def _start_ocr(
        self,
        images: list[Image.Image],
        lang: str,
        engine: str,
        on_done: Optional[Callable[[], None]] = None,
    ) -> None:
        diac_enabled = self.diacritics_cb.isChecked() if hasattr(self, "diacritics_cb") else False
        self._diacritics_failed = False
        self._diacritics_enabled_cache = diac_enabled
        self.ocr_thread = create_ocr_thread(engine, images, lang, diacritics_enabled=diac_enabled)
        self.ocr_thread.finished.connect(self._ocr_finished)
        self.ocr_thread.ocr_results.connect(self._set_ocr_results)
        if hasattr(self.ocr_thread, "diacritics_failed"):
            self.ocr_thread.diacritics_failed.connect(self._on_diacritics_failed)
        if hasattr(self.ocr_thread, "failed"):
            self.ocr_thread.failed.connect(self._on_easyocr_thread_failed)
        if on_done:
            self.ocr_thread.finished.connect(on_done)

        if engine == "EasyOCR":
            self.progress_dialog = QProgressDialog(
                "Načítám EasyOCR model...", "Zrušit", 0, 0, self
            )
            if hasattr(self.ocr_thread, "model_loading"):
                self.ocr_thread.model_loading.connect(self._on_model_loaded)
        else:
            self.progress_dialog = QProgressDialog(
                "Probíhá OCR (Tesseract)...", "Zrušit", 0, 100, self
            )

        self.ocr_thread.progress.connect(self._on_ocr_progress)
        self.progress_dialog.show()
        self.ocr_thread.start()

    def _get_alternate_engine(self, engine: str) -> str:
        return "EasyOCR" if engine == "Tesseract" else "Tesseract"

    def _build_attempts(self) -> list[tuple[str, str, list[Image.Image]]]:
        # Zachováno pro kompatibilitu / explicitní retry pokud bude potřeba,
        # ale automatický fallback již není používán. Vrací pouze jednorázový
        # pokus s aktuálním nastavením uživatele.
        base_engine = self.engine_combo.currentText()
        base_lang = self.lang_combo.currentText()
        return [(base_engine, base_lang, list(self.scanned_images))]

    def _run_ocr_once(
        self, lang: str, engine: str, images: list[Image.Image]
    ) -> None:
        loop = QEventLoop()
        self._start_ocr(images, lang, engine, on_done=loop.quit)
        loop.exec()

    def _run_single_attempt(
        self, engine: str, lang: str, images: list[Image.Image]
    ) -> bool:
        """Spustí JEDEN OCR pokus a vrátí True pokud byl úspěšný (platný text)."""
        if self._scan_cancelled:
            return False
        self._ocr_mode = "macro"
        self._run_ocr_once(lang, engine, images)
        return self._has_ocr_text()

    def _run_ocr_flow(self, interactive: bool) -> None:
        """Nové chování: první úspěšný výsledek stačí. Žádný automatický fallback.

        1. Uživatel zvolí engine/lang → jeden pokus.
        2. Pokud vrátí platný text → považováno za dokončené, nabídnout retry pouze explicitně.
        3. Pokud selže → nabídnout volby včetně „Zkusit jiný engine“ (explicitně).
        """
        base_engine = self.engine_combo.currentText()
        base_lang = self.lang_combo.currentText()
        # Guard znovu – pokud je EasyOCR a není ready, již jsme deferovali v run_ocr,
        # ale pro přímé volání (pending save, retry) zkontroluj znovu
        if base_engine == "EasyOCR":
            if not self._handle_easyocr_guard_before_ocr(base_engine, base_lang, {"mode": "full", "interactive": interactive}):
                return
        images = list(self.scanned_images)
        found = self._run_single_attempt(base_engine, base_lang, images)

        if interactive:
            if found:
                self._ocr_mode = "interactive"
                used_engine = base_engine
                # Krátké oznámení – engine, ne technické detaily
                speak(f"OCR dokončeno, použit {used_engine}.")
                # pending save nesmí zobrazit druhý náhled
                if not getattr(self, "_pending_save_format", None):
                    self._show_ocr_preview(self._last_text)
            else:
                action = self._ask_no_text_action()
                if action == "retry_other":
                    self._retry_with_alternate_engine(interactive=True)
                elif action == "rescan":
                    speak("Znovu skenuji stránky a opakuji OCR.")
                    self.scan_pages()
                    if self.scanned_images:
                        self._run_ocr_flow(interactive=True)
                elif action == "continue":
                    self._ocr_mode = "interactive"
                    if not getattr(self, "_pending_save_format", None):
                        self._show_ocr_preview(self._last_text)

    def _retry_with_alternate_engine(self, interactive: bool = True) -> None:
        """Jednorázový explicitní pokus druhým enginem. Nemění engine_combo.

        Zachová původní výsledek do úspěchu nového pokusu. Při selhání obnoví původní.
        """
        base_engine = self.engine_combo.currentText()
        alt_engine = self._get_alternate_engine(base_engine)
        lang = self.lang_combo.currentText()
        images = list(self.scanned_images)
        # Snapshot pro případ selhání
        backup_results = list(self.last_ocr_results)
        backup_text = self._last_text
        backup_pages = self._last_edited_pages
        backup_original = self._last_original_text
        backup_total = getattr(self, "_ocr_total_pages", 0)
        backup_announced = getattr(self, "_ocr_last_announced_page", 0)

        speak(f"Zkouším jiný OCR engine {alt_engine}.")
        # Dočasně přepnout počitadla pro progress
        self._ocr_total_pages = len(images)
        self._ocr_last_announced_page = 0
        self._run_ocr_once(lang, alt_engine, images)
        if not self._has_ocr_text():
            # Obnovit původní výsledek
            self.last_ocr_results = backup_results
            self._last_text = backup_text
            self._last_edited_pages = backup_pages
            self._last_original_text = backup_original
            self._ocr_total_pages = backup_total
            self._ocr_last_announced_page = backup_announced
            QMessageBox.warning(self, "OCR bez výsledku", f"Pokus s {alt_engine} neprodukoval žádný text. Původní výsledek byl zachován.")
            speak("Nový pokus selhal, původní výsledek zachován.")
            if interactive and not getattr(self, "_pending_save_format", None):
                # Nabídnout znovu náhled původního výsledku pokud existoval
                if backup_text and backup_text.strip():
                    self._show_ocr_preview(self._last_text)
            return
        # Úspěch – zachovat nový výsledek (engine_combo se nemění)
        speak(f"OCR dokončeno, použit {alt_engine}.")
        if interactive and not getattr(self, "_pending_save_format", None):
            self._ocr_mode = "interactive"
            self._show_ocr_preview(self._last_text)

    def _retry_incremental_with_alternate_engine(self, start_idx: int) -> None:
        """Jednorázový retry pouze pro nově přidané stránky (start_idx..). Nemění engine_combo."""
        new_images = self.scanned_images[start_idx:]
        if not new_images:
            return
        base_engine = self.engine_combo.currentText()
        alt_engine = self._get_alternate_engine(base_engine)
        lang = self.lang_combo.currentText()
        # Snapshot před retry
        backup_results = list(self.last_ocr_results)
        backup_text = self._last_text
        backup_pages = self._last_edited_pages
        backup_original = self._last_original_text

        speak(f"Zkouším jiný OCR engine {alt_engine} pro nové stránky.")
        # Provést jednorázový pokus na new_images
        new_results_holder: list[list[OcrResult]] = []
        new_text_holder: list[str] = []
        loop = QEventLoop()

        def on_results(r):
            new_results_holder.clear()
            new_results_holder.extend(r)

        def on_finished(txt: str):
            new_text_holder.append(txt)
            loop.quit()

        diac_enabled = self.diacritics_cb.isChecked() if hasattr(self, "diacritics_cb") else False
        self._diacritics_failed = False
        self._diacritics_enabled_cache = diac_enabled
        # Dočasné hodnoty pro progress
        prev_total = getattr(self, "_ocr_total_pages", 0)
        prev_announced = getattr(self, "_ocr_last_announced_page", 0)
        self._ocr_total_pages = len(self.scanned_images)
        self._ocr_last_announced_page = start_idx
        self.ocr_thread = create_ocr_thread(alt_engine, new_images, lang, diacritics_enabled=diac_enabled)
        self.ocr_thread.ocr_results.connect(on_results)
        self.ocr_thread.finished.connect(on_finished)
        if hasattr(self.ocr_thread, "diacritics_failed"):
            self.ocr_thread.diacritics_failed.connect(self._on_diacritics_failed)
        if alt_engine == "EasyOCR":
            self.progress_dialog = QProgressDialog("Načítám EasyOCR model...", "Zrušit", 0, 0, self)
            self.ocr_thread.model_loading.connect(self._on_model_loaded)
        else:
            self.progress_dialog = QProgressDialog("Probíhá OCR (Tesseract)...", "Zrušit", 0, 100, self)
        self.ocr_thread.progress.connect(self._on_ocr_progress)
        self.progress_dialog.show()
        self.ocr_thread.start()
        loop.exec()
        self.progress_dialog.cancel()

        # Kontrola úspěchu pro nové stránky
        has_text = False
        for page in new_results_holder:
            for item in page:
                if item.text and item.text.strip():
                    has_text = True
                    break
            if has_text:
                break
        # Fallback přes text pokud results prázdné ale full_text něco obsahuje
        if not has_text and new_text_holder and new_text_holder[0].strip():
            stripped = re.sub(r"---\s*Stránka\s+\d+\s*---", "", new_text_holder[0]).strip()
            has_text = bool(stripped)

        if not has_text or not new_results_holder:
            # Obnovit (incremental ještě nic nemergoval – stačí vrátit backup)
            self.last_ocr_results = backup_results
            self._last_text = backup_text
            self._last_edited_pages = backup_pages
            self._last_original_text = backup_original
            self._ocr_total_pages = prev_total
            self._ocr_last_announced_page = prev_announced
            QMessageBox.warning(self, "OCR bez výsledku", f"Pokus s {alt_engine} pro nové stránky neprodukoval text. Původní výsledek zachován.")
            speak("Nový pokus selhal, původní výsledek zachován.")
            # Zobrazit původní náhled pokud existoval
            if backup_text and backup_text.strip():
                self._show_ocr_preview(self._last_text, is_incremental=True, incremental_start_idx=start_idx)
            return
        # Úspěch – merge stejně jako v _run_ocr_incremental
        old_results = backup_results[:start_idx] if len(backup_results) >= start_idx else backup_results
        # Pokud backup byl prázdný, všechny jsou nové
        if not old_results and start_idx == 0:
            old_results = []
        # Přegeneruj corrected_new_text s posunutým číslováním
        corrected_new_text = ""
        for i, page_results in enumerate(new_results_holder):
            page_num = start_idx + i + 1
            page_text_parts = [r.display_text for r in page_results]
            page_text = "\n".join(page_text_parts)
            corrected_new_text += f"--- Stránka {page_num} ---\n{page_text}\n\n"
        # Merge results
        self.last_ocr_results = old_results + new_results_holder
        # Merge text
        old_text_for_merge = backup_text if backup_text else ""
        # Pokud starý text existoval, připoj nové; jinak použij corrected
        # Pro případ kdy old_text neobsahuje nové stránky, rekonstruuj správně
        if start_idx == 0:
            self._last_text = corrected_new_text
        else:
            # Ořízni backup na start_idx stránek pokud byl již kompletní
            if backup_text and backup_text.strip():
                self._last_text = old_text_for_merge.rstrip() + "\n\n" + corrected_new_text if old_text_for_merge.strip() else corrected_new_text
            else:
                self._last_text = corrected_new_text
        try:
            pages = _split_text_by_page_headers(self._last_text)
            if pages is not None and len(pages) == len(self.last_ocr_results):
                self._last_edited_pages = pages
            else:
                self._last_edited_pages = None
        except Exception:
            self._last_edited_pages = None
        try:
            orig_parts = []
            for idx, page in enumerate(self.last_ocr_results):
                if page:
                    texts = [r.original_text if r.original_text is not None else r.text for r in page]
                    orig_parts.append(f"--- Stránka {idx + 1} ---\n" + "\n".join(texts) + "\n\n")
            self._last_original_text = "".join(orig_parts) if any(orig_parts) else self._last_text
        except Exception:
            self._last_original_text = self._last_text
        self.btn_read.setEnabled(True)
        speak(f"OCR dokončeno, použit {alt_engine}.")
        if not getattr(self, "_pending_save_format", None):
            self._show_ocr_preview(self._last_text, is_incremental=True, incremental_start_idx=start_idx)

    def _ask_no_text_action(self) -> str:
        speak("Nebyl rozpoznán žádný text.")
        base_engine = self.engine_combo.currentText()
        alt_engine = self._get_alternate_engine(base_engine)
        msg = QMessageBox(self)
        msg.setWindowTitle("Nerozpoznán žádný text")
        msg.setText(f"Zvolený engine {base_engine} nerozpoznal žádný text.")
        msg.setInformativeText("Co chcete udělat?")
        btn_retry = msg.addButton(f"Zkusit jiný OCR engine ({alt_engine})", QMessageBox.ButtonRole.ActionRole)
        btn_retry.setAccessibleName(f"Zkusit jiný OCR engine {alt_engine}")
        btn_retry.setAccessibleDescription("Spustí jednorázový pokus druhým enginem. Původní výsledek zůstane zachován.")
        btn_rescan = msg.addButton("Znovu naskenovat", QMessageBox.ButtonRole.YesRole)
        btn_rescan.setAccessibleName("Znovu naskenovat stránky")
        btn_cont = msg.addButton("Pokračovat k náhledu", QMessageBox.ButtonRole.ActionRole)
        btn_cont.setAccessibleName("Pokračovat k náhledu a uložení")
        btn_cancel = msg.addButton("Zrušit", QMessageBox.ButtonRole.NoRole)
        btn_cancel.setAccessibleName("Zrušit")
        msg.exec()
        clicked = msg.clickedButton()
        if clicked == btn_retry:
            return "retry_other"
        if clicked == btn_rescan:
            return "rescan"
        if clicked == btn_cont:
            return "continue"
        return "cancel"

    def _execute_ocr(self) -> None:
        # Makro / neinteraktivní cesta – také single-attempt, bez automatického fallbacku
        self._ocr_total_pages = len(self.scanned_images)
        self._ocr_last_announced_page = 0
        base_engine = self.engine_combo.currentText()
        base_lang = self.lang_combo.currentText()
        # Jednorázový pokus, výsledek v last_ocr_results / _last_text via _ocr_finished
        self._run_single_attempt(base_engine, base_lang, list(self.scanned_images))

    def _on_ocr_progress(self, value: int) -> None:
        self.progress_dialog.setValue(value)
        # Per-page voice feedback – throttled to once per page
        total = getattr(self, "_ocr_total_pages", len(self.scanned_images)) or 1
        # Derive current page from progress percent
        current_page = int(round(value / 100 * total)) if total else 0
        if current_page > 0 and current_page != getattr(self, "_ocr_last_announced_page", 0):
            # Avoid double-announce at 100 which is handled separately
            if value != 100:
                speak(f"OCR stránky {current_page} z {total}.")
            self._ocr_last_announced_page = current_page
        if value == 100:
            speak("OCR dokončeno")

    def _on_model_loaded(self, value: int) -> None:
        if value == 100:
            self.progress_dialog.setMaximum(100)
            self.progress_dialog.setLabelText("Probíhá OCR (EasyOCR)...")
            speak("EasyOCR model načten, zahajuji rozpoznávání")

    def _set_ocr_results(self, results: list[list[OcrResult]]) -> None:
        self.last_ocr_results = results

    def _on_diacritics_failed(self, msg: str) -> None:
        self._diacritics_failed = True

    def _ocr_finished(self, full_text: str) -> None:
        self.progress_dialog.cancel()
        self._last_text = full_text
        # Inicializuj per-page z full_text
        try:
            pages = _split_text_by_page_headers(full_text)
            if pages is not None and len(pages) == len(self.scanned_images):
                self._last_edited_pages = pages
            else:
                self._last_edited_pages = None
        except Exception:
            self._last_edited_pages = None
        # Rekonstrukce originálu z OcrResult.original_text pro zachování původního výsledku
        try:
            orig_parts = []
            for idx, page in enumerate(self.last_ocr_results):
                # pokud máme bbox výsledky, poskládej originál, jinak použij full_text
                if page:
                    texts = [r.original_text if r.original_text is not None else r.text for r in page]
                    orig_parts.append(f"--- Stránka {idx + 1} ---\n" + "\n".join(texts) + "\n\n")
                else:
                    orig_parts.append("")
            self._last_original_text = "".join(orig_parts) if any(orig_parts) else full_text
        except Exception:
            self._last_original_text = full_text
        self.btn_read.setEnabled(True)
        total = getattr(self, "_ocr_total_pages", len(self.scanned_images))
        # Hlasová odezva dle diakritiky
        from diacritics import should_diacritize
        lang = self.lang_combo.currentText() if hasattr(self, "lang_combo") else ""
        diac_active = should_diacritize(lang, self._diacritics_enabled_cache)
        if self._diacritics_failed and diac_active:
            speak("OCR dokončeno. Oprava české diakritiky se nepodařila. Byl zachován původní text.")
        elif diac_active:
            speak("OCR dokončeno a česká diakritika opravena.")
            if total:
                speak(f"OCR dokončeno pro {total} stránek.")
        else:
            if total:
                speak(f"OCR dokončeno pro {total} stránek.")
        # Pokud běží pending save (OCR spuštěno z dialogu Uložit), nepřekrývej náhledem
        if getattr(self, "_pending_save_format", None):
            return
        if self._ocr_mode == "interactive":
            self._show_ocr_preview(full_text)

    def _show_ocr_preview(self, full_text: str, is_incremental: bool = False, incremental_start_idx: int = 0) -> None:
        """Zobrazí náhled OCR textu k editaci – SAMOSTATNÝ krok OCR, bez ukládání.

        Obsahuje explicitní volbu „Zkusit jiný OCR engine“ (jednorázový pokus, nemění engine_combo).
        Retry se týká celého dokumentu, u inkrementálního pouze nově přidaných stránek.
        """
        speak("Zobrazuji náhled rozpoznaného textu. Můžete jej upravit nebo přečíst. Uložení provedete tlačítkem Uložit dokument.")
        alt_engine = self._get_alternate_engine(self.engine_combo.currentText()) if hasattr(self, "engine_combo") else ""
        dialog = OcrPreviewDialog(full_text, self, alt_engine_name=alt_engine)
        result = dialog.exec()
        if result == OcrPreviewDialog.RESULT_RETRY:
            # Explicitní žádost – jednorázový pokus druhým enginem, původní výsledek zachován do úspěchu
            if is_incremental:
                self._retry_incremental_with_alternate_engine(incremental_start_idx)
            else:
                self._retry_with_alternate_engine(interactive=True)
            return
        if result != QDialog.DialogCode.Accepted:
            self.btn_read.setFocus()
            return
        edited_text = dialog.get_text()
        self._last_text = edited_text
        try:
            pages = _split_text_by_page_headers(edited_text)
            if pages is not None:
                self._last_edited_pages = pages
            else:
                if len(self.scanned_images) == 1:
                    self._last_edited_pages = [edited_text]
                else:
                    self._last_edited_pages = None
        except Exception:
            self._last_edited_pages = None
        # Po náhledu nabídni uložení přes samostatné tlačítko, ale ne automaticky
        # (uživatel stiskne Uložit dokument)
        self.btn_save_document.setFocus()
        speak("Náhled uložen. Pro uložení stiskněte Uložit dokument.")

    def _ocr_show_save_ui(self, full_text: str) -> None:
        """Legacy wrapper – zachován pro makra. Nově jen zobrazí náhled bez auto-uložení."""
        self._show_ocr_preview(full_text)
        # Původní auto-uložení odstraněno – oddělení OCR a ukládání.

    # ---------- Uložit dokument – dvoustupňový tok ----------
    def save_document(self) -> None:
        """Hlavní vstup pro tlačítko Uložit dokument – nikdy nespouští OCR automaticky."""
        if not self.scanned_images:
            QMessageBox.warning(self, "Chyba", "Nejdříve naskenujte stránky.")
            speak("Žádné stránky k uložení.")
            return
        speak("Otevírám dialog Co chcete uložit.")
        kind_dlg = ChooseSaveKindDialog(self)
        if kind_dlg.exec() != QDialog.DialogCode.Accepted:
            self.page_list.setFocus()
            return
        kind = kind_dlg.selected_kind()
        if kind == "images":
            self._save_images_only_flow()
        else:  # ocr
            fmt_dlg = ChooseOcrFormatDialog(self)
            if fmt_dlg.exec() != QDialog.DialogCode.Accepted:
                self.page_list.setFocus()
                return
            fmt = fmt_dlg.selected_format()  # txt / docx / pdf_ocr
            self._pending_save_format = fmt
            try:
                self._try_save_with_ocr(fmt)
            finally:
                # pokud nebyl spuštěn OCR, vyčisti hned; pokud byl pending save úspěšný, už je vyčištěn
                if self._pending_save_format == fmt:
                    # nebyl spuštěn OCR, nebo selhal – vyčisti
                    self._pending_save_format = None

    def _save_images_only_flow(self) -> None:
        """PDF bez OCR – používá přímo self.scanned_images, nikdy OCR."""
        speak("Ukládám PDF bez OCR – pouze obrázky.")
        default_dir = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DocumentsLocation)
        if not default_dir:
            default_dir = os.path.expanduser("~")
        default_path = os.path.join(default_dir, "naskenovany_dokument.pdf")
        path, _ = QFileDialog.getSaveFileName(self, "Uložit dokument", default_path, "PDF – naskenované stránky (*.pdf)")
        if not path:
            self.page_list.setFocus()
            return
        if not path.lower().endswith(".pdf"):
            base, ext = os.path.splitext(path)
            if ext.lower() not in (".txt", ".pdf", ".docx"):
                path = path + ".pdf"
            elif ext.lower() != ".pdf":
                path = base + ".pdf"
        if os.path.exists(path):
            if not self._confirm_overwrite(path):
                speak("Ukládání zrušeno.")
                self.page_list.setFocus()
                return
        try:
            self._save_pdf_without_ocr(path)
        except Exception as e:
            QMessageBox.critical(self, "Chyba při ukládání", f"Nepodařilo se uložit soubor:\n{e}")
            speak("Chyba při ukládání souboru.")
            self.page_list.setFocus()

    def _try_save_with_ocr(self, fmt: str) -> None:
        total = len(self.scanned_images)
        missing = self._missing_ocr_count()
        is_empty = self._is_ocr_empty()
        is_complete = self._is_ocr_complete()
        # fmt: txt, docx, pdf_ocr
        if is_complete:
            self._do_save_ocr_format(fmt)
            self._pending_save_format = None
            return
        # parciální nebo prázdné – zobraz dialog
        if is_empty:
            action = self._ask_ocr_not_available(fmt)
        else:
            action = self._ask_partial_ocr(missing, total, fmt)
        if action == "run_ocr":
            # Spustit OCR a po dokončení automaticky pokračovat
            self._run_ocr_for_pending_save(fmt)
        elif action == "save_without" and fmt == "pdf_ocr":
            # Uživatel chce uložit bez OCR místo s OCR
            self._pending_save_format = None
            self._save_images_only_flow()
        elif action == "cancel":
            self._pending_save_format = None
            speak("Ukládání zrušeno.")
            self.page_list.setFocus()
        else:
            # pro txt/docx není save_without – jen cancel/run
            self._pending_save_format = None
            speak("Ukládání zrušeno.")
            self.page_list.setFocus()

    def _ask_ocr_not_available(self, fmt: str) -> str:
        """Dialog OCR není k dispozici. Vrátí run_ocr / save_without / cancel."""
        speak("OCR není k dispozici.")
        fmt_human = {"txt": "TXT", "docx": "DOCX", "pdf_ocr": "PDF s OCR"}[fmt]
        msg = QMessageBox(self)
        msg.setWindowTitle("OCR není k dispozici")
        if fmt == "pdf_ocr":
            msg.setText("Pro vytvoření PDF s rozpoznaným textem je nejprve potřeba provést OCR.")
        elif fmt == "txt":
            msg.setText("Textový obsah zatím není k dispozici, protože nebylo provedeno OCR.")
        else:
            msg.setText("Obsah pro Word dokument zatím není k dispozici, protože nebylo provedeno OCR.")
        msg.setInformativeText(f"Požadovaný formát: {fmt_human}. Co chcete udělat?")
        btn_run = msg.addButton("Spustit OCR", QMessageBox.ButtonRole.YesRole)
        btn_run.setAccessibleName("Spustit OCR a poté uložit")
        btn_run.setAccessibleDescription("Explicitně spustí OCR a po dokončení automaticky uloží dokument.")
        if fmt == "pdf_ocr":
            btn_without = msg.addButton("Uložit PDF bez OCR", QMessageBox.ButtonRole.ActionRole)
            btn_without.setAccessibleName("Uložit PDF bez OCR – pouze obrázky")
        else:
            btn_without = None
        btn_cancel = msg.addButton("Zrušit", QMessageBox.ButtonRole.NoRole)
        btn_cancel.setAccessibleName("Zrušit ukládání")
        msg.setDefaultButton(btn_run)
        # Přístupnost: fokus na Spustit OCR
        msg.exec()
        clicked = msg.clickedButton()
        if clicked == btn_run:
            return "run_ocr"
        if btn_without is not None and clicked == btn_without:
            return "save_without"
        return "cancel"

    def _ask_partial_ocr(self, missing: int, total: int, fmt: str) -> str:
        speak("Některé stránky nemají OCR.")
        msg = QMessageBox(self)
        msg.setWindowTitle("Chybějící OCR")
        msg.setText(f"Některé stránky ještě nemají OCR ({total-missing} z {total} má OCR, chybí {missing}). Chcete nejprve spustit OCR pro chybějící stránky?")
        fmt_human = {"txt": "TXT", "docx": "DOCX", "pdf_ocr": "PDF s OCR"}[fmt]
        msg.setInformativeText(f"Požadovaný formát: {fmt_human}. Chybějící stránky lze rozpoznat doplňkovým OCR.")
        btn_run = msg.addButton("Spustit OCR", QMessageBox.ButtonRole.YesRole)
        btn_run.setAccessibleName("Spustit OCR pro chybějící stránky a poté uložit")
        if fmt == "pdf_ocr":
            btn_without = msg.addButton("Uložit bez OCR", QMessageBox.ButtonRole.ActionRole)
            btn_without.setAccessibleName("Uložit všechny stránky bez OCR")
        else:
            btn_without = None
        btn_cancel = msg.addButton("Zrušit", QMessageBox.ButtonRole.NoRole)
        btn_cancel.setAccessibleName("Zrušit ukládání")
        msg.setDefaultButton(btn_run)
        msg.exec()
        clicked = msg.clickedButton()
        if clicked == btn_run:
            return "run_ocr"
        if btn_without is not None and clicked == btn_without:
            return "save_without"
        return "cancel"

    def _run_ocr_for_pending_save(self, fmt: str) -> None:
        """Spustí OCR explicitně a po úspěchu automaticky uloží pending formát."""
        # zachovej fmt v self._pending_save_format (už nastaveno)
        total_before = len(self.scanned_images)
        try:
            if self._is_ocr_empty():
                # full OCR
                self._ocr_total_pages = total_before
                self._ocr_last_announced_page = 0
                speak(f"Spouštím OCR pro uložení {fmt}, celkem {total_before} stránek.")
                self._run_ocr_flow(interactive=False)
                # _run_ocr_flow v non-interactive neukazuje preview, ale nastaví _last_text
                # pokud našlo text, pokračuj k uložení
            else:
                # incremental – doplnit chybějící
                missing_start = len(self.last_ocr_results)
                self._run_ocr_incremental(missing_start)
                # _run_ocr_incremental už vrací pokud pending, bez preview
        except Exception as e:
            QMessageBox.critical(self, "Chyba OCR", f"OCR selhalo:\n{e}")
            speak("OCR selhalo.")
            self._pending_save_format = None
            return
        # Ověř že OCR nyní kompletní
        if self._is_ocr_empty() or not self._has_ocr_text():
            QMessageBox.warning(self, "OCR bez výsledku", "OCR bylo dokončeno, ale nebyl rozpoznán žádný text. Dokument nebude uložen.")
            speak("OCR neprodukovalo text.")
            self._pending_save_format = None
            return
        # Automaticky pokračovat v původně zvoleném ukládání
        try:
            self._do_save_ocr_format(fmt)
        finally:
            self._pending_save_format = None

    def _do_save_ocr_format(self, fmt: str) -> None:
        """Uloží OCR formát – voláno pouze když OCR existuje, nikdy nespouští OCR."""
        # vyber filtr a příponu
        filters = {"txt": "Textový dokument (*.txt)", "docx": "Word dokument (*.docx)", "pdf_ocr": "PDF s rozpoznaným textem (*.pdf)"}
        suffixes = {"txt": ".txt", "docx": ".docx", "pdf_ocr": ".pdf"}
        chosen_filter = filters[fmt]
        suffix = suffixes[fmt]
        default_dir = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DocumentsLocation)
        if not default_dir:
            default_dir = os.path.expanduser("~")
        name_map = {"txt": "naskenovany_dokument.txt", "docx": "naskenovany_dokument.docx", "pdf_ocr": "naskenovany_dokument.pdf"}
        default_path = os.path.join(default_dir, name_map[fmt])
        speak("Vyberte umístění pro uložení souboru.")
        path, _ = QFileDialog.getSaveFileName(self, "Uložit dokument", default_path, chosen_filter)
        if not path:
            self.page_list.setFocus()
            return
        if not path.lower().endswith(suffix):
            base, ext = os.path.splitext(path)
            if ext.lower() not in (".txt", ".pdf", ".docx"):
                path = path + suffix
            elif ext.lower() != suffix:
                path = base + suffix
        if os.path.exists(path):
            if not self._confirm_overwrite(path):
                speak("Ukládání zrušeno.")
                self.page_list.setFocus()
                return
        try:
            if fmt == "pdf_ocr":
                self._save_pdf_with_ocr(path)
            elif fmt == "docx":
                self._save_docx_pages(path)
            else:
                # pro TXT použij _last_text (již obsahuje headery)
                self._save_txt_pages(path, self._last_text)
        except Exception as e:
            QMessageBox.critical(self, "Chyba při ukládání", f"Nepodařilo se uložit soubor:\n{e}")
            speak("Chyba při ukládání souboru.")
            self.page_list.setFocus()

    def _confirm_overwrite(self, path: str) -> bool:
        msg = QMessageBox(self)
        msg.setWindowTitle("Soubor již existuje")
        msg.setText(f"Soubor již existuje:\n{path}")
        msg.setInformativeText("Chcete jej přepsat?")
        btn_over = msg.addButton("Přepsat", QMessageBox.ButtonRole.YesRole)
        btn_over.setAccessibleName("Ano, přepsat soubor")
        btn_cancel = msg.addButton("Zrušit", QMessageBox.ButtonRole.NoRole)
        btn_cancel.setAccessibleName("Zrušit, neukládat")
        msg.setDefaultButton(btn_cancel)
        msg.exec()
        return msg.clickedButton() == btn_over

    def read_last_text(self) -> None:
        # Při zapnuté opravě čte processed_text (už v _last_text), jinak originál
        text_to_read = self._last_text
        if self._last_text:
            speak(text_to_read)
        else:
            speak("Není žádný text k přečtení.")

    # ---------- Export ----------
    def export_images(self) -> None:
        if not self.scanned_images:
            QMessageBox.warning(self, "Chyba", "Nejdříve naskenujte stránky.")
            return

        formats = [("JPEG", ".jpg"), ("PNG", ".png"), ("TIFF", ".tif")]
        fmt_items = [f"{name} (*{ext})" for name, ext in formats]
        fmt_combo = QComboBox()
        fmt_combo.addItems(fmt_items)
        fmt_combo.setAccessibleName("Formát obrázků")

        msg = QMessageBox(self)
        msg.setWindowTitle("Export obrázků")
        msg.setText("Zvolte formát obrázků:")
        msg.setInformativeText("Poté budete vyzváni k výběru cílové složky.")
        msg.layout().addWidget(fmt_combo)
        msg.addButton(QMessageBox.StandardButton.Ok)
        msg.addButton(QMessageBox.StandardButton.Cancel)

        if msg.exec() != QMessageBox.StandardButton.Ok:
            return

        dir_path = QFileDialog.getExistingDirectory(self, "Vyberte cílovou složku")
        if not dir_path:
            return

        sel_idx = fmt_combo.currentIndex()
        _, ext = formats[sel_idx]
        pil_fmt = formats[sel_idx][0]

        for i, img in enumerate(self.scanned_images):
            fname = os.path.join(dir_path, f"stranka_{i+1:03d}{ext}")
            if pil_fmt == "JPEG":
                img = img.convert("RGB")
                img.save(fname, pil_fmt, quality=95)
            else:
                img.save(fname, pil_fmt)

        QMessageBox.information(
            self, "Hotovo",
            f"{len(self.scanned_images)} obrázků uloženo do {dir_path}"
        )
        speak("Obrázky exportovány.")

    # ---------- Import PDF se selektivním OCR ----------
    def import_pdfs(self) -> None:
        """GUI workflow importu PDF. Tezka prace bezi ve workeru, dialogy
        a speak() pouze zde v GUI vlakne."""
        default_dir = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DocumentsLocation)
        if not default_dir:
            default_dir = os.path.expanduser("~")
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Importovat PDF", default_dir, "PDF (*.pdf)")
        if not paths:
            return
        if len(paths) == 1:
            speak("Byl vybrán 1 PDF soubor.")
        elif 2 <= len(paths) <= 4:
            speak(f"Byly vybrány {len(paths)} PDF soubory.")
        else:
            speak(f"Bylo vybráno {len(paths)} PDF souborů.")

        # Rychla analyza (cista textova extrakce, ms) - souhrn pro dialog.
        analyses = [pdf_import.analyze_pdf(p) for p in paths]
        summary = "\n\n".join(pdf_import.format_analysis_summary(a) for a in analyses)
        need_total = sum(len(a.pages_needing_ocr) for a in analyses if a.ok)

        if need_total == 0:
            speak("Všechny vybrané dokumenty již obsahují použitelný text. OCR není potřeba.")
            QMessageBox.information(
                self, "Import PDF",
                summary + "\n\nVšechny vybrané dokumenty již obsahují použitelný text. OCR není potřeba.")
            self.btn_import_pdf.setFocus()
            return

        speak(summary)
        if PdfImportSummaryDialog(summary, self).exec() != QDialog.DialogCode.Accepted:
            speak("Import zrušen.")
            self.btn_import_pdf.setFocus()
            return

        info = (f"OCR proběhne pouze na {need_total} "
                f"stránkách bez použitelného textu. Původní soubory zůstanou zachovány.")
        settings_dlg = ImportOcrSettingsDialog(
            self.engine_combo.currentText(), self.lang_combo.currentIndex(), info, self)
        if settings_dlg.exec() != QDialog.DialogCode.Accepted:
            speak("Import zrušen.")
            self.btn_import_pdf.setFocus()
            return

        # Propsat volbu zpet do stavajiciho systemu nastaveni (zadny druhy system).
        engine = settings_dlg.selected_engine()
        lang_idx = settings_dlg.selected_lang_index()
        eng_idx = self.engine_combo.findText(engine)
        if eng_idx >= 0:
            self.engine_combo.setCurrentIndex(eng_idx)
        if 0 <= lang_idx < self.lang_combo.count():
            self.lang_combo.setCurrentIndex(lang_idx)
        try:
            self._save_settings()
        except Exception:
            pass
        lang_raw = self.lang_combo.currentText()
        diac_enabled = self.diacritics_cb.isChecked() if hasattr(self, "diacritics_cb") else False

        # Predbezna kontrola konfliktu nazvu - VZDY v GUI vlakne.
        overwrite_allowed: dict[str, bool] = {}
        for analysis in analyses:
            if not analysis.ok or not analysis.pages_needing_ocr:
                continue
            out = pdf_import.default_output_path(analysis.path)
            if os.path.exists(out):
                if self._confirm_overwrite(out):
                    overwrite_allowed[out] = True
                else:
                    overwrite_allowed[out] = False
        jobs, skipped = pdf_import.build_jobs(
            analyses, engine, lang_raw, diac_enabled, overwrite_allowed)
        if not jobs:
            speak("Žádný soubor k zpracování. Import ukončen.")
            detail = "\n".join(s.message for s in skipped) or "Ukládání bylo zrušeno."
            QMessageBox.information(self, "Import PDF",
                                    f"Žádný soubor k zpracování.\n{detail}")
            self.btn_import_pdf.setFocus()
            return

        self._run_pdf_import_jobs(jobs, skipped)

    def _run_pdf_import_jobs(self, jobs: list, skipped: list) -> None:
        """Spusti davkovy worker a po dokonceni zobrazi pristupny souhrn."""
        self._pdf_import_job_info: dict[str, tuple[int, int, int]] = {}
        for i, job in enumerate(jobs):
            self._pdf_import_job_info[job.src_path] = (i + 1, len(jobs), len(job.pages_to_ocr))
        total_units = sum(len(j.pages_to_ocr) for j in jobs) or 1

        self._pdf_import_worker = PdfImportWorker(jobs)
        self._pdf_import_results: list = []
        progress = QProgressDialog("Zahajuji import PDF...", "Zrušit", 0, total_units, self)
        progress.setWindowTitle("Import PDF")
        progress.setAccessibleName("Průběh importu PDF")
        self._pdf_import_progress = progress
        worker = self._pdf_import_worker
        worker.job_started.connect(self._on_pdf_import_job_started)
        worker.job_message.connect(self._on_pdf_import_job_message)
        worker.overall_progress.connect(self._on_pdf_import_overall)
        worker.batch_finished.connect(self._on_pdf_import_batch_finished)

        loop = QEventLoop()
        worker.batch_finished.connect(loop.quit)
        progress.canceled.connect(self._cancel_pdf_import)
        worker.start()
        progress.show()
        loop.exec()
        try:
            progress.cancel()
        except Exception:
            pass

        results = list(self._pdf_import_results) + list(skipped)
        ok = [r for r in results if r.status == "ok"]
        failed = [r for r in results if r.status == "failed"]
        cancelled = [r for r in results if r.status == "cancelled"]
        skipped_only = [r for r in results if r.status == "skipped"]

        lines = [f"Import dokončen: úspěšně {len(ok)}, "
                 f"přeskočeno {len(skipped_only)}, "
                 f"selhalo {len(failed)}, zrušeno {len(cancelled)}."]
        for r in ok:
            lines.append(f"OK: {os.path.basename(r.src_path)} → {os.path.basename(r.out_path)}")
        for r in failed:
            lines.append(f"Selhalo: {os.path.basename(r.src_path)} – {r.message}")
        for r in skipped_only:
            lines.append(f"Přeskočeno: {os.path.basename(r.src_path)} – {r.message}")
        text = "\n".join(lines)
        if failed or cancelled:
            QMessageBox.warning(self, "Import PDF", text)
        else:
            QMessageBox.information(self, "Import PDF", text)
        if ok and not failed and not cancelled:
            speak(f"OCR dokončeno. Zpracováno {len(ok)} dokumentů.")
        else:
            speak(f"Import dokončen. Úspěšně {len(ok)}, selhalo {len(failed)}, "
                  f"přeskočeno {len(skipped_only)}, zrušeno {len(cancelled)}.")
        self.btn_import_pdf.setFocus()
        self._pdf_import_worker = None

    def _update_pdf_import_label(self, text: str) -> None:
        try:
            if getattr(self, "_pdf_import_progress", None) is not None:
                self._pdf_import_progress.setLabelText(text)
        except Exception:
            pass

    def _on_pdf_import_job_started(self, src_path: str, idx: int, total: int) -> None:
        base = os.path.basename(src_path)
        info = getattr(self, "_pdf_import_job_info", {}).get(src_path)
        n = f", celkem {info[2]} stran k OCR" if info else ""
        self._update_pdf_import_label(f"Zpracovávám soubor {idx} ze {total}: {base}{n}.")
        speak(f"Zpracovávám soubor {idx} ze {total}: {base}.")

    def _on_pdf_import_job_message(self, src_path: str, text: str) -> None:
        base = os.path.basename(src_path)
        info = getattr(self, "_pdf_import_job_info", {}).get(src_path)
        if info is not None:
            idx, total, _n = info
            self._update_pdf_import_label(
                f"Zpracovávám soubor {idx} ze {total}: {base}, {text}.")
        else:
            self._update_pdf_import_label(f"{base}: {text}.")

    def _on_pdf_import_overall(self, done: int, total: int) -> None:
        try:
            if getattr(self, "_pdf_import_progress", None) is not None:
                self._pdf_import_progress.setMaximum(max(int(total), 1))
                self._pdf_import_progress.setValue(int(done))
        except Exception:
            pass

    def _on_pdf_import_batch_finished(self, results: list) -> None:
        self._pdf_import_results = list(results)

    def _cancel_pdf_import(self) -> None:
        speak("Ruším import PDF.")
        try:
            if getattr(self, "_pdf_import_worker", None) is not None:
                self._pdf_import_worker.requestInterruption()
        except Exception:
            pass

    # ---------- Save helpers ----------
    def _get_docx_pages(self) -> list[str]:
        """Vrátí list textů per-page pro DOCX podle skutečných stránek.

        Priorita:
        1) _last_edited_pages pokud délka == scanned_images
        2) rekonstrukce z last_ocr_results (display_text)
        3) fallback rozdělení _last_text
        """
        total = len(self.scanned_images)
        # 1) upravené stránky pokud sedí počet
        if self._last_edited_pages is not None and len(self._last_edited_pages) == total:
            return list(self._last_edited_pages)
        # Zkusit re-split _last_text pokud obsahuje headery a sedí
        if self._last_text:
            pages = _split_text_by_page_headers(self._last_text)
            if pages is not None and len(pages) == total:
                return pages
        # 2) rekonstrukce z OCR výsledků
        if self.last_ocr_results and len(self.last_ocr_results) == total:
            out: list[str] = []
            for page in self.last_ocr_results:
                if not page:
                    out.append("")
                else:
                    # Spoj display_text jednotlivých OcrResult s novým řádkem
                    # (zachová odřádkování, ne přidává umělé \n\n mezi každým slovem)
                    texts = [r.display_text for r in page]
                    # Pokud je více výsledků, spoj je novým řádkem – DOCX pak splitne \n\n na odstavce
                    out.append("\n".join(texts) if len(texts) > 1 else (texts[0] if texts else ""))
            return out
        # 3) fallback – rozděl _last_text i když nesedí, nebo vrať jako jednu stránku
        if self._last_text:
            pages = _split_text_by_page_headers(self._last_text)
            if pages is not None:
                # Doplň/zkrát na total
                if len(pages) < total:
                    pages = pages + [""] * (total - len(pages))
                return pages[:total]
            return [self._last_text]
        return [""] * total if total else []

    def _save_txt_pages(self, path: str, edited_text: str) -> None:
        """Uloží TXT s oddělovači --- Stránka N --- v pořadí 1..N, čitelné pro NVDA."""
        total = len(self.scanned_images)
        # Pokud edited_text již obsahuje headery a počet sedí nebo je alespoň 1, ulož přímo
        pages = _split_text_by_page_headers(edited_text)
        if pages is not None and len(pages) == total:
            # Obsah již má správné headery – ulož edited_text přímo (zachová přesnou editaci)
            # Ale normalizuj aby každý header byl přesně "--- Stránka N ---"
            # Pro jednoduchost ulož přímo edited_text pokud obsahuje headery
            content = edited_text
            # Zajisti, že soubor končí newline
            if not content.endswith("\n"):
                content += "\n"
        elif total > 0:
            # Generuj per-page z edited_text per-page pokud sedí, jinak fallback na edited split nebo OCR
            if pages is not None and len(pages) == total:
                content = ""
                for i, p in enumerate(pages):
                    content += f"--- Stránka {i+1} ---\n{p}\n\n"
            elif self._last_edited_pages is not None and len(self._last_edited_pages) == total:
                content = ""
                for i, p in enumerate(self._last_edited_pages):
                    content += f"--- Stránka {i+1} ---\n{p}\n\n"
            else:
                # Pokud edited_text nemá headery ale máme total, zkusit rozdělitEdited nebo použít celý edited_text jako stránku 1 + prázdné?
                # Nejbezpečnější: pokud edited_text neobsahuje headery a je jen jeden dokument,
                # ulož ho jako souvislý text s headery podle skutečných stránek.
                # Pokud edited_text obsahuje více odstavců ale bez headerů, nelze bezpečně rekonstruovat per-page,
                # takže pokud total==1 ulož přímo, jinak vygeneruj z OCR nebo z edited_text jako celek pro stránku 1.
                if total == 1:
                    # Pokud jedna stránka, bez headeru je ok, ale pro konzistenci přidej header
                    # Pokud uživatel explicitně smazal header, respektuj jeho editaci – ulož bez headeru
                    # Detekce: pokud edited_text nemá header a total==1, ulož edited_text přímo
                    content = edited_text
                else:
                    # Více stránek ale text bez headerů – pokus se generovat z OCR per-page pokud dostupné
                    if self.last_ocr_results and len(self.last_ocr_results) == total:
                        # Pokud je edited_text výrazně odlišný od OCR, může být uživatelská editace
                        # – v tom případě bez spolehlivého per-page rozdělení je bezpečnější uložit
                        # edited_text jako celek bez umělého dělení.
                        # Preferuj edited_text jako celek pro TXT čitelnost, ale zachovej informaci o stránkách:
                        # Zkusíme: pokud edited_text obsahuje "\n\n" ale ne headery, uložíme ho s headery pouze pro první stránku?
                        # Ne – specifikace říká TXT může obsahovat souvislý text všech stránek s oddělovači.
                        # Pokud nemáme per-page, použij OCR per-page jako zdroj a ignoruj edited? To by zahodilo editaci.
                        # Proto: ulož edited_text jako souvislý text s fallback generováním oddělovačů jen pokud edited_text je per-page.
                        # Aktuálně: edited_text bez headerů pro více stran → ulož edited_text přímo (bez umělého dělení) + přidej info?
                        # Rozhodnutí: ulož edited_text přímo, protože umělé dělení by bylo nebezpečné.
                        # Ale pro konzistenci a NVDA čitelnost je lepší mít headery. Pokud edited_text nemá headery,
                        # vytvoř per-page z OCR a nepoužívej edited – dokument by ztratil editaci. To je horší.
                        # Kompromis: pokud edited_text bez headerů a total>1, ulož s headery kde stránka 1 = edited_text, ostatní prázdné?
                        # Ne, to je matoucí. Nejmenší překvapení: ulož edited_text přímo.
                        content = edited_text
                    else:
                        content = edited_text
        else:
            content = edited_text

        with open(path, "w", encoding="utf-8") as f:
            f.write(content)
        self._show_save_success(path, "Textový dokument (.txt)", "Textový dokument")

    def _save_pdf_without_ocr(self, path: str) -> None:
        """Uloží stránky jako obrázky. OCR nebude proveden. Použije přímo self.scanned_images."""
        doc = fitz.open()
        dpi = self.dpi_combo.currentData()
        if dpi is None:
            dpi = 300
        for img in self.scanned_images:
            page = doc.new_page(width=img.width / dpi * 72, height=img.height / dpi * 72)
            tmp_path = None
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
                    tmp_path = tmp.name
                    img.save(tmp_path, "JPEG", quality=95)
                page.insert_image(page.rect, filename=tmp_path)
            finally:
                if tmp_path and os.path.exists(tmp_path):
                    try:
                        os.remove(tmp_path)
                    except OSError:
                        pass
        doc.save(path)
        self._show_save_success(path, "PDF – naskenované stránky (.pdf)", "PDF bez OCR")

    def _save_pdf_with_ocr(self, path: str) -> None:
        """Zachová vzhled naskenovaných stránek a přidá textovou vrstvu z OCR. Použije self.last_ocr_results, nespouští OCR."""
        dpi = self.dpi_combo.currentData()
        if dpi is None:
            dpi = 300
        # Veškerá logika vrstvy je v testovatelné build_searchable_pdf().
        # Záměrně nemění zdroj textu (last_ocr_results, ne ruční editace) –
        # viz ticket: editace z OcrPreviewDialog je samostatný problém.
        stats = build_searchable_pdf(list(self.scanned_images),
                                     list(self.last_ocr_results),
                                     int(dpi), path)
        logger.info("PDF s OCR uloženo: %s stats=%s", path, stats)
        self._show_save_success(path, "PDF s rozpoznaným textem (.pdf)", "PDF s OCR")

    def _save_pdf(self, path: str) -> None:
        """Legacy alias – pro kompatibilitu volá _save_pdf_with_ocr pokud existuje OCR, jinak bez."""
        if self._is_ocr_complete():
            self._save_pdf_with_ocr(path)
        else:
            # zachovej původní chování – vlož OCR vrstvu pokud existuje částečně
            # ale nový tok by měl volat explicitně jednu z variant
            if self.last_ocr_results:
                self._save_pdf_with_ocr(path)
            else:
                self._save_pdf_without_ocr(path)

    def _save_docx(self, text: str, path: str) -> None:
        """Legacy wrapper – zachován pro kompatibilitu (makra). Nově volá per-page logiku."""
        # Aktualizuj _last_text a _last_edited_pages pro per-page logiku
        self._last_text = text
        pages = _split_text_by_page_headers(text)
        if pages is not None:
            self._last_edited_pages = pages
        self._save_docx_pages(path)

    def _save_docx_pages(self, path: str) -> None:
        """Uloží DOCX rozdělený podle skutečných naskenovaných stránek (page break pouze mezi stránkami)."""
        doc = Document()
        pages = self._get_docx_pages()
        total = len(pages)
        for idx, page_text in enumerate(pages):
            # Rozděl na odstavce podle dvojitého odřádkování, ale bez přidávání page break mezi odstavci
            if page_text is None:
                page_text = ""
            # Zachovej prázdné stránky
            if not page_text.strip():
                doc.add_paragraph("")
            else:
                paragraphs = page_text.split("\n\n")
                for para in paragraphs:
                    # Prázdné odstavce zachovej jako prázdný paragraph
                    doc.add_paragraph(para)
            if idx < total - 1:
                doc.add_page_break()
        doc.save(path)
        self._show_save_success(path, "Word dokument (.docx)", "Word dokument")

    def _show_save_success(self, path: str, format_human: str, format_short: str) -> None:
        """Společný úspěšný dialog – přístupný, s cestou v textu, krátkou hlasovou hláškou."""
        QMessageBox.information(
            self, "Hotovo",
            f"Dokument byl úspěšně uložen.\nFormát: {format_human}\nCesta: {path}"
        )
        # Hlasově krátce, bez celé cesty
        speak(f"Dokument byl úspěšně uložen. Dokument byl uložen jako {format_short}.")
        self.btn_scan.setFocus()

    # ---------- Macros ----------
    def _rebuild_macro_buttons(self) -> None:
        while self._macro_buttons_layout.count():
            item = self._macro_buttons_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        for macro in self.macro_manager.macros:
            container = QWidget()
            row = QHBoxLayout(container)
            row.setContentsMargins(0, 0, 0, 0)
            label = macro.name
            if macro.shortcut:
                label += f" ({macro.shortcut})"
            btn_run = QPushButton(label)
            btn_run.setAccessibleName(f"Spustit makro: {macro.name}")
            btn_run.clicked.connect(lambda checked, m=macro: self._run_macro(m))
            row.addWidget(btn_run)

            btn_edit = QPushButton("Upravit")
            btn_edit.setAccessibleName(f"Upravit makro: {macro.name}")
            btn_edit.clicked.connect(lambda checked, m=macro: self._open_macro_editor(m))
            row.addWidget(btn_edit)

            btn_rename = QPushButton("Přejmenovat")
            btn_rename.setAccessibleName(f"Přejmenovat makro: {macro.name}")
            btn_rename.clicked.connect(lambda checked, m=macro: self._rename_macro(m))
            row.addWidget(btn_rename)

            btn_delete = QPushButton("Smazat")
            btn_delete.setAccessibleName(f"Smazat makro: {macro.name}")
            btn_delete.clicked.connect(lambda checked, m=macro: self._delete_macro(m))
            row.addWidget(btn_delete)

            self._macro_buttons_layout.addWidget(container)

    def _open_macro_editor(self, macro: Optional[Macro] = None) -> None:
        dialog = MacroEditorDialog(self, self.macro_manager, macro)
        dialog.exec()

    def _run_macro(self, macro: Macro) -> None:
        self.speak(f"Spouštím makro: {macro.name}")
        self._macro_runner.run(macro)
        self.speak("Makro dokončeno.")

    def _rename_macro(self, macro: Macro) -> None:
        new_name, ok = QInputDialog.getText(
            self, "Přejmenovat makro", "Nový název makra:",
            text=macro.name,
        )
        if ok and new_name.strip():
            self.macro_manager.rename(macro, new_name.strip())
            self.macro_manager.load_all()
            self._rebuild_macro_buttons()
            self._setup_macro_shortcuts()
            self.speak(f"Makro přejmenováno na {new_name.strip()}")

    def _delete_macro(self, macro: Macro) -> None:
        msg = QMessageBox(self)
        msg.setWindowTitle("Smazat makro")
        msg.setText(f"Opravdu chcete smazat makro {macro.name}?")
        msg.setInformativeText("Tuto akci nelze vrátit.")
        btn_yes = msg.addButton("Ano", QMessageBox.ButtonRole.YesRole)
        btn_yes.setAccessibleName(f"Ano, smazat makro {macro.name}")
        btn_no = msg.addButton("Ne", QMessageBox.ButtonRole.NoRole)
        btn_no.setAccessibleName("Ne, ponechat makro")
        msg.exec()
        if msg.clickedButton() == btn_yes:
            self.macro_manager.delete(macro)
            self._rebuild_macro_buttons()
            self._setup_macro_shortcuts()
            self.speak(f"Makro {macro.name} smazáno.")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    font = app.font()
    font.setPointSize(11)
    app.setFont(font)
    win = ScanApp()
    win.show()
    sys.exit(app.exec())
