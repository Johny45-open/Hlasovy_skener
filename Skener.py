import sys
import io
import threading
import os
import tempfile
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel,
    QComboBox, QSpinBox, QMessageBox, QFileDialog, QDialog,
    QProgressDialog
)
from PyQt6.QtGui import QPixmap
from PyQt6.QtCore import Qt, QEventLoop
from PIL import Image, ImageQt
from docx import Document
import fitz
from gtts import gTTS
import pygame
from accessible_output2.outputs.auto import Auto

# Importy z vlastních modulů
from scanner_engine import NAPS2Scanner, ScanThread
from ocr_engine import create_ocr_thread

# ------------------ Hlasový výstup ------------------
speaker = Auto()

def speak(text, lang="cs"):
    """Spustí hlasový výstup přes systémový screen reader."""
    speaker.output(text)
# ------------------ Náhled dialog ------------------
class PreviewDialog(QDialog):
    def __init__(self, pil_img):
        super().__init__()
        self.setWindowTitle("Náhled skenu")
        self.image = pil_img

        self.label = QLabel()
        self.update_image()

        btn_rotate = QPushButton("Otočit")
        btn_rotate.clicked.connect(self.rotate_image)

        btn_ok = QPushButton("Potvrdit náhled")
        btn_ok.clicked.connect(self.accept)

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(btn_rotate)
        layout.addWidget(btn_ok)
        self.setLayout(layout)

    def update_image(self):
        qimg = ImageQt.ImageQt(self.image)
        pix = QPixmap.fromImage(qimg).scaled(500, 700, Qt.AspectRatioMode.KeepAspectRatio)
        self.label.setPixmap(pix)

    def rotate_image(self):
        self.image = self.image.rotate(-90, expand=True)
        self.update_image()

# ------------------ Hlavní aplikace ------------------
class ScanApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("NAPS2 skener + OCR + hlas")
        self.scanner = NAPS2Scanner()
        self.scanned_images = []
        self.last_ocr_results = []

        layout = QVBoxLayout()

        self.device_combo = QComboBox()
        self.reload_devices()

        self.source_combo = QComboBox()
        self.source_combo.addItems(["Sklo", "Podavač"])

        self.dpi_spin = QSpinBox()
        self.dpi_spin.setRange(100, 1200)
        self.dpi_spin.setValue(300)

        self.color_combo = QComboBox()
        self.color_combo.addItems(["Barevný", "Šedý", "ČB"])

        self.lang_combo = QComboBox()
        self.lang_combo.setAccessibleName("Jazyk OCR")
        self.lang_combo.addItems([
            "ces (Čeština)", "eng (Angličtina)", "deu (Němčina)",
            "fra (Francouzština)", "ita (Italština)", "pol (Polština)"
        ])

        layout.addWidget(QLabel("Jazyk OCR:"))
        layout.addWidget(self.lang_combo)

        self.engine_combo = QComboBox()
        self.engine_combo.addItems(["Tesseract", "EasyOCR"])
        self.engine_combo.setAccessibleName("OCR engine")
        layout.addWidget(QLabel("OCR engine:"))
        layout.addWidget(self.engine_combo)

        layout.addWidget(QLabel("Vyber profil NAPS2:"))
        self.device_combo.setAccessibleName("Vyber profil NAPS2")
        layout.addWidget(self.device_combo)
        layout.addWidget(QLabel("Zdroj papíru:"))
        self.source_combo.setAccessibleName("Zdroj papíru")
        layout.addWidget(self.source_combo)
        layout.addWidget(QLabel("Rozlišení DPI:"))
        self.dpi_spin.setAccessibleName("Rozlišení DPI")
        layout.addWidget(self.dpi_spin)
        layout.addWidget(QLabel("Režim:"))
        self.color_combo.setAccessibleName("Režim barev")
        layout.addWidget(self.color_combo)
        
        self.btn_scan = QPushButton("Skenovat stránku")
        self.btn_scan.setAccessibleName("Skenovat stránku")
        self.btn_scan.clicked.connect(self.scan_pages)
        layout.addWidget(self.btn_scan)

        self.btn_ocr = QPushButton("Spustit OCR a uložit")
        self.btn_ocr.setAccessibleName("Spustit OCR a uložit")
        self.btn_ocr.clicked.connect(self.run_ocr)
        self.btn_ocr.setEnabled(False)
        layout.addWidget(self.btn_ocr)

        self.setLayout(layout)

    def reload_devices(self):
        try:
            devices = self.scanner.list_devices()
            self.device_combo.clear()
            for d in devices:
                self.device_combo.addItem(d)
            if devices:
                self.scanner.connect_device(devices[0])
            elif not self.scanner.naps2_exe:
                QMessageBox.warning(self, "NAPS2 nenalezen", "NAPS2.Console.exe nebyl nalezen. Ujistěte se, že je NAPS2 nainstalován.")
        except Exception as e:
            QMessageBox.critical(self, "Chyba", f"Nenašel jsem žádný skener:\n{e}")

    def scan_pages(self):
        speak("Zahajuji skenování.")
        self.scanned_images.clear()
        while True:
            if not self.scan_page_single():
                break
            
            # Vlastní dialog pro lokalizaci tlačítek
            msg = QMessageBox(self)
            msg.setWindowTitle("Další")
            msg.setText("Další stránka?")
            btn_yes = msg.addButton("Ano", QMessageBox.ButtonRole.YesRole)
            btn_no = msg.addButton("Ne", QMessageBox.ButtonRole.NoRole)
            msg.exec()
            
            if msg.clickedButton() != btn_yes:
                break
        
        speak("Skenování dokončeno.")
        if self.scanned_images:
            self.btn_ocr.setEnabled(True)

    def scan_page_single(self):
        device_name = self.device_combo.currentText()
        if not device_name:
            QMessageBox.warning(self, "Chyba", "Není vybrán žádný skener.")
            return False
            
        self.scanner.connect_device(device_name)

        dpi = self.dpi_spin.value()
        source = self.source_combo.currentText()
        color_text = self.color_combo.currentText()
        # NAPS2 používá: Color, Grayscale, BlackWhite
        color_mode = "Color" if color_text == "Barevný" else "Grayscale" if color_text == "Šedý" else "BlackWhite"

        self.progress_dialog = QProgressDialog("Skenuji (přes NAPS2)...", "Zrušit", 0, 0, self)
        self.progress_dialog.show()

        self.scan_thread = ScanThread(self.scanner, dpi, color_mode, source)
        self.scan_thread.finished.connect(self.scan_finished)
        self.scan_thread.start()

        loop = QEventLoop()
        self.scan_thread.finished.connect(loop.quit)
        self.scan_thread.finished.connect(self.progress_dialog.cancel)
        loop.exec()

        return hasattr(self, 'last_scanned_image') and self.last_scanned_image

    def scan_finished(self, img):
        self.progress_dialog.cancel()
        preview = PreviewDialog(img)
        if preview.exec():
            self.scanned_images.append(preview.image)
            self.last_scanned_image = True
        else:
            self.last_scanned_image = False

    def run_ocr(self):
        speak("Zahajuji OCR.")
        lang = self.lang_combo.currentText()
        engine = self.engine_combo.currentText()

        self.ocr_thread = create_ocr_thread(engine, self.scanned_images, lang)
        self.ocr_thread.finished.connect(self.ocr_finished)
        self.ocr_thread.ocr_results.connect(self.set_ocr_results)

        if engine == "EasyOCR":
            self.progress_dialog = QProgressDialog("Načítám EasyOCR model...", "Zrušit", 0, 0, self)
            self.ocr_thread.model_loading.connect(self._on_model_loaded)
        else:
            self.progress_dialog = QProgressDialog("Probíhá OCR (Tesseract)...", "Zrušit", 0, 100, self)
            self.ocr_thread.progress.connect(self.progress_dialog.setValue)

        self.progress_dialog.show()
        self.ocr_thread.start()

    def _on_model_loaded(self, value):
        if value == 100:
            self.progress_dialog.setMaximum(100)
            self.progress_dialog.setLabelText("Probíhá OCR (EasyOCR)...")
            self.ocr_thread.progress.connect(self.progress_dialog.setValue)

    def set_ocr_results(self, results):
        self.last_ocr_results = results

    def ocr_finished(self, full_text):
        self.progress_dialog.cancel()
        speak("OCR dokončeno. Vyberte umístění pro uložení souboru.")
        path, _ = QFileDialog.getSaveFileName(self, "Uložit", "", "Text (*.txt);;PDF (*.pdf);;Word (*.docx)")
        if not path: return

        if path.lower().endswith(".pdf"):
            self.save_pdf(path)
        elif path.lower().endswith(".docx"):
            self.save_docx(full_text, path)
        else:
            with open(path, "w", encoding="utf-8") as f: f.write(full_text)
            QMessageBox.information(self, "Hotovo", "Uloženo.")
            speak("Soubor byl úspěšně uložen.")

    def save_pdf(self, path):
        doc = fitz.open()
        dpi = self.dpi_spin.value()
        for i, img in enumerate(self.scanned_images):
            # Vytvoření stránky
            page = doc.new_page(width=img.width / dpi * 72, height=img.height / dpi * 72)
            
            # Vložení obrázku
            with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
                img.save(tmp.name, "JPEG", quality=85)
                page.insert_image(page.rect, filename=tmp.name)
            
            # Vložení skrytého textu pro vyhledávání
            for item in self.last_ocr_results[i]:
                if 'bbox' in item:
                    bbox = item['bbox']
                    x0, y0 = bbox[0][0], bbox[0][1]
                    x1, y1 = bbox[2][0], bbox[2][1]
                    rect = fitz.Rect(x0 / dpi * 72, y0 / dpi * 72, x1 / dpi * 72, y1 / dpi * 72)
                    page.insert_textbox(rect, item['text'], fontsize=0, fill_opacity=0)
        
        doc.save(path)
        QMessageBox.information(self, "Hotovo", "PDF uloženo.")

    def save_docx(self, text, path):
        doc = Document()
        for part in text.split("\n\n"):
            doc.add_paragraph(part)
            doc.add_page_break()
        doc.save(path)
        QMessageBox.information(self, "Hotovo", "DOCX uloženo.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = ScanApp()
    win.show()
    sys.exit(app.exec())
