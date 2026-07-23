from __future__ import annotations
import sys
import os
import tempfile
from typing import Optional

from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QComboBox, QSpinBox, QMessageBox, QFileDialog, QDialog, QListWidget,
    QProgressDialog, QTextEdit, QCheckBox, QAbstractItemView, QGroupBox
)
from PyQt6.QtGui import QPixmap, QShortcut, QKeySequence
from PyQt6.QtCore import Qt, QEventLoop, QSettings, QTimer
from PIL import Image, ImageQt, ImageFilter, ImageOps
from docx import Document
import fitz
from accessible_output2.outputs.auto import Auto

from scanner_engine import NAPS2Scanner, ScanThread
from ocr_engine import create_ocr_thread, OcrResult

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
    def __init__(self, text: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Náhled OCR textu")
        self.setMinimumSize(600, 400)

        self.text_edit = QTextEdit()
        self.text_edit.setPlainText(text)
        self.text_edit.setAccessibleName("OCR text k editaci")

        btn_layout = QHBoxLayout()
        btn_read = QPushButton("Přečíst text")
        btn_read.setAccessibleName("Přečíst text hlasem")
        btn_read.clicked.connect(self.read_text)
        btn_layout.addWidget(btn_read)

        btn_save = QPushButton("Uložit")
        btn_save.setAccessibleName("Uložit a pokračovat")
        btn_save.setDefault(True)
        btn_save.clicked.connect(self.accept)
        btn_layout.addWidget(btn_save)

        btn_cancel = QPushButton("Zrušit")
        btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(btn_cancel)

        layout = QVBoxLayout()
        layout.addWidget(self.text_edit)
        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def read_text(self) -> None:
        speak(self.text_edit.toPlainText())

    def showEvent(self, event) -> None:
        super().showEvent(event)
        self.text_edit.setFocus()

    def get_text(self) -> str:
        return self.text_edit.toPlainText()

# ------------------ Hlavní aplikace ------------------
class ScanApp(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Hlasový skener – NAPS2 + OCR")
        self.setMinimumSize(500, 600)

        self.settings = QSettings("HlasovySkener", "HlasovySkener")
        self.scanner = NAPS2Scanner()
        self.scanned_images: list[Image.Image] = []
        self.last_ocr_results: list[list[OcrResult]] = []
        self.last_scanned_image: Optional[bool] = None
        self._scan_cancelled = False

        self._build_ui()
        self._setup_shortcuts()
        self._load_settings()

    # ---------- UI ----------
    def _build_ui(self) -> None:
        layout = QVBoxLayout()

        # -- Nastavení OCR --
        ocr_group = QGroupBox("OCR")
        ocr_group.setAccessibleName("Nastavení OCR")
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

        ocr_group.setLayout(ocr_layout)
        layout.addWidget(ocr_group)

        # -- Nastavení skeneru --
        scan_group = QGroupBox("Skener")
        scan_group.setAccessibleName("Nastavení skeneru")
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
        self.dpi_spin = QSpinBox()
        self.dpi_spin.setRange(100, 1200)
        self.dpi_spin.setValue(300)
        self.dpi_spin.setAccessibleName("Rozlišení DPI")
        lbl_dpi.setBuddy(self.dpi_spin)
        scan_layout.addWidget(lbl_dpi)
        scan_layout.addWidget(self.dpi_spin)

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

        # -- Tlačítka OCR --
        ocr_btn_layout = QHBoxLayout()
        self.btn_ocr = QPushButton("OCR a uložit (Ctrl+O)")
        self.btn_ocr.setAccessibleName("Spustit OCR a uložit")
        self.btn_ocr.clicked.connect(self.run_ocr)
        self.btn_ocr.setEnabled(False)
        ocr_btn_layout.addWidget(self.btn_ocr)

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

        self.setLayout(layout)

        # -- Tab order --
        self.setTabOrder(self.lang_combo, self.engine_combo)
        self.setTabOrder(self.engine_combo, self.preprocess_cb)
        self.setTabOrder(self.preprocess_cb, self.batch_cb)
        self.setTabOrder(self.batch_cb, self.device_combo)
        self.setTabOrder(self.device_combo, self.source_combo)
        self.setTabOrder(self.source_combo, self.dpi_spin)
        self.setTabOrder(self.dpi_spin, self.color_combo)
        self.setTabOrder(self.color_combo, self.btn_scan)
        self.setTabOrder(self.btn_scan, self.btn_scan_all)
        self.setTabOrder(self.btn_scan_all, self.page_list)
        self.setTabOrder(self.page_list, btn_delete_page)
        self.setTabOrder(btn_delete_page, btn_clear_pages)
        self.setTabOrder(btn_clear_pages, self.btn_ocr)
        self.setTabOrder(self.btn_ocr, self.btn_read)
        self.setTabOrder(self.btn_read, btn_export_img)

    def _setup_shortcuts(self) -> None:
        QShortcut(QKeySequence("Ctrl+S"), self).activated.connect(self.scan_pages)
        QShortcut(QKeySequence("Ctrl+Shift+S"), self).activated.connect(self.scan_all_pages)
        QShortcut(QKeySequence("Ctrl+O"), self).activated.connect(self.run_ocr)
        QShortcut(QKeySequence("Ctrl+P"), self).activated.connect(self.read_last_text)
        QShortcut(QKeySequence("Ctrl+E"), self).activated.connect(self.export_images)
        QShortcut(QKeySequence("Ctrl+Q"), self).activated.connect(self.close)
        QShortcut(QKeySequence("Delete"), self).activated.connect(self.delete_page)

    # ---------- Settings persistence ----------
    def _load_settings(self) -> None:
        lang_idx = self.settings.value("ocr/lang_index", 0, type=int)
        if 0 <= lang_idx < self.lang_combo.count():
            self.lang_combo.setCurrentIndex(lang_idx)
        engine_idx = self.settings.value("ocr/engine_index", 0, type=int)
        if 0 <= engine_idx < self.engine_combo.count():
            self.engine_combo.setCurrentIndex(engine_idx)
        dpi_val = self.settings.value("scan/dpi", 300, type=int)
        self.dpi_spin.setValue(dpi_val)
        source_idx = self.settings.value("scan/source_index", 0, type=int)
        if 0 <= source_idx < self.source_combo.count():
            self.source_combo.setCurrentIndex(source_idx)
        color_idx = self.settings.value("scan/color_index", 0, type=int)
        if 0 <= color_idx < self.color_combo.count():
            self.color_combo.setCurrentIndex(color_idx)
        self.batch_cb.setChecked(self.settings.value("scan/batch", False, type=bool))
        self.preprocess_cb.setChecked(self.settings.value("ocr/preprocess", True, type=bool))

    def _save_settings(self) -> None:
        self.settings.setValue("ocr/lang_index", self.lang_combo.currentIndex())
        self.settings.setValue("ocr/engine_index", self.engine_combo.currentIndex())
        self.settings.setValue("scan/dpi", self.dpi_spin.value())
        self.settings.setValue("scan/source_index", self.source_combo.currentIndex())
        self.settings.setValue("scan/color_index", self.color_combo.currentIndex())
        self.settings.setValue("scan/batch", self.batch_cb.isChecked())
        self.settings.setValue("ocr/preprocess", self.preprocess_cb.isChecked())

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
        speak("Zahajuji skenování.")
        self.scanned_images.clear()
        self.page_list.clear()
        self.last_ocr_results.clear()
        self.btn_ocr.setEnabled(False)
        self.btn_read.setEnabled(False)

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
            self.page_list.addItems([f"Stránka {i+1}" for i in range(len(self.scanned_images))])
            self.page_list.setFocus()
            speak(f"Skenování dokončeno. {len(self.scanned_images)} stránek.")
        else:
            self.btn_scan.setFocus()
            speak("Skenování dokončeno, žádné stránky.")

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
        dpi = self.dpi_spin.value()
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
            processed = preview.image
            if self.preprocess_cb.isChecked():
                processed = preprocess_image(processed)
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
        self.page_list.takeItem(row)
        del self.scanned_images[row]
        if self.last_ocr_results and row < len(self.last_ocr_results):
            del self.last_ocr_results[row]
        remaining = len(self.scanned_images)
        if remaining == 0:
            self.btn_ocr.setEnabled(False)
            self.btn_read.setEnabled(False)
            speak("Všechny stránky smazány.")
        else:
            speak(f"Stránka smazána. Zbývá {remaining} stránek.")
            # Focus na stejnou pozici nebo poslední
            next_row = min(row, remaining - 1)
            self.page_list.setCurrentRow(next_row)
            self.page_list.setFocus()

    def clear_pages(self) -> None:
        self.page_list.clear()
        self.scanned_images.clear()
        self.last_ocr_results.clear()
        self.btn_ocr.setEnabled(False)
        self.btn_read.setEnabled(False)

    # ---------- OCR ----------
    def run_ocr(self) -> None:
        if not self.scanned_images:
            QMessageBox.warning(self, "Chyba", "Nejdříve naskenujte stránky.")
            return

        speak("Zahajuji OCR.")
        lang = self.lang_combo.currentText()
        engine = self.engine_combo.currentText()

        self.ocr_thread = create_ocr_thread(engine, self.scanned_images, lang)
        self.ocr_thread.finished.connect(self._ocr_finished)
        self.ocr_thread.ocr_results.connect(self._set_ocr_results)

        if engine == "EasyOCR":
            self.progress_dialog = QProgressDialog(
                "Načítám EasyOCR model...", "Zrušit", 0, 0, self
            )
            self.ocr_thread.model_loading.connect(self._on_model_loaded)
        else:
            self.progress_dialog = QProgressDialog(
                "Probíhá OCR (Tesseract)...", "Zrušit", 0, 100, self
            )

        self.ocr_thread.progress.connect(self._on_ocr_progress)
        self.progress_dialog.show()
        self.ocr_thread.start()

    def _on_ocr_progress(self, value: int) -> None:
        self.progress_dialog.setValue(value)
        if value in (25, 50, 75):
            speak(f"OCR z {value} procent hotovo")
        elif value == 100:
            speak("OCR dokončeno")

    def _on_model_loaded(self, value: int) -> None:
        if value == 100:
            self.progress_dialog.setMaximum(100)
            self.progress_dialog.setLabelText("Probíhá OCR (EasyOCR)...")
            speak("EasyOCR model načten, zahajuji rozpoznávání")

    def _set_ocr_results(self, results: list[list[OcrResult]]) -> None:
        self.last_ocr_results = results

    def _ocr_finished(self, full_text: str) -> None:
        self.progress_dialog.cancel()

        speak("Zobrazuji náhled rozpoznaného textu. Můžete jej upravit, přečíst nebo uložit.")
        dialog = OcrPreviewDialog(full_text, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        edited_text = dialog.get_text()
        self._last_text = edited_text
        self.btn_read.setEnabled(True)

        speak("Vyberte umístění pro uložení souboru.")
        path, selected_filter = QFileDialog.getSaveFileName(
            self, "Uložit", "",
            "Text (*.txt);;PDF (*.pdf);;Word (*.docx)"
        )
        if not path:
            self.btn_read.setFocus()
            return

        if path.lower().endswith(".pdf"):
            self._save_pdf(path)
        elif path.lower().endswith(".docx"):
            self._save_docx(edited_text, path)
        else:
            with open(path, "w", encoding="utf-8") as f:
                f.write(edited_text)
            QMessageBox.information(self, "Hotovo", "Text uložen.")
            speak("Soubor byl úspěšně uložen.")
            self.btn_scan.setFocus()

    def read_last_text(self) -> None:
        text = getattr(self, '_last_text', None)
        if text:
            speak(text)

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

    # ---------- Save helpers ----------
    def _save_pdf(self, path: str) -> None:
        doc = fitz.open()
        dpi = self.dpi_spin.value()
        for i, img in enumerate(self.scanned_images):
            page = doc.new_page(
                width=img.width / dpi * 72,
                height=img.height / dpi * 72
            )

            tmp_path = None
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
                    tmp_path = tmp.name
                    img.save(tmp_path, "JPEG", quality=95)
                page.insert_image(page.rect, filename=tmp_path)

                if i < len(self.last_ocr_results):
                    for item in self.last_ocr_results[i]:
                        if item.bbox is not None:
                            bbox = item.bbox
                            x0, y0 = bbox[0][0], bbox[0][1]
                            x1, y1 = bbox[2][0], bbox[2][1]
                            rect = fitz.Rect(
                                x0 / dpi * 72, y0 / dpi * 72,
                                x1 / dpi * 72, y1 / dpi * 72
                            )
                            page.insert_textbox(rect, item.text, fontsize=0, fill_opacity=0)
            finally:
                if tmp_path and os.path.exists(tmp_path):
                    try:
                        os.remove(tmp_path)
                    except OSError:
                        pass

        doc.save(path)
        QMessageBox.information(self, "Hotovo", "PDF uloženo.")
        speak("PDF soubor uložen.")
        self.btn_scan.setFocus()

    def _save_docx(self, text: str, path: str) -> None:
        doc = Document()
        for part in text.split("\n\n"):
            doc.add_paragraph(part)
            doc.add_page_break()
        doc.save(path)
        QMessageBox.information(self, "Hotovo", "DOCX uloženo.")
        speak("Soubor uložen.")
        self.btn_scan.setFocus()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = ScanApp()
    win.show()
    sys.exit(app.exec())
