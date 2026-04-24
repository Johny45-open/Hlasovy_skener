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

from scanner_engine import WIAScanner, ScanThread
from ocr_engine import OCRThread

# ------------------ Hlasový výstup ------------------
def speak(text, lang="cs"):
    """Spustí hlasový výstup ve vlákně."""
    def _speak():
        try:
            tts = gTTS(text=text, lang=lang)
            mp3_fp = io.BytesIO()
            tts.write_to_fp(mp3_fp)
            mp3_fp.seek(0)
            pygame.mixer.init()
            pygame.mixer.music.load(mp3_fp, "mp3")
            pygame.mixer.music.play()
            while pygame.mixer.music.get_busy():
                pygame.time.Clock().tick(10)
            pygame.mixer.quit()
        except Exception as e:
            print("Chyba při hlasovém výstupu:", e)
    threading.Thread(target=_speak, daemon=True).start()

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
        self.setWindowTitle("WIA skener + OCR + hlas")
        self.scanner = WIAScanner()
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
        self.lang_combo.addItems([
            "ces (Čeština)", "eng (Angličtina)", "deu (Němčina)",
            "fra (Francouzština)", "ita (Italština)", "pol (Polština)"
        ])

        layout.addWidget(QLabel("Jazyk OCR:"))
        layout.addWidget(self.lang_combo)
        layout.addWidget(QLabel("Vyber skener:"))
        layout.addWidget(self.device_combo)
        layout.addWidget(QLabel("Zdroj papíru:"))
        layout.addWidget(self.source_combo)
        layout.addWidget(QLabel("Rozlišení DPI:"))
        layout.addWidget(self.dpi_spin)
        layout.addWidget(QLabel("Režim:"))
        layout.addWidget(self.color_combo)
        
        self.btn_scan = QPushButton("Skenovat stránku")
        self.btn_scan.clicked.connect(self.scan_pages)
        layout.addWidget(self.btn_scan)

        self.btn_ocr = QPushButton("Spustit OCR a uložit")
        self.btn_ocr.clicked.connect(self.run_ocr)
        self.btn_ocr.setEnabled(False)
        layout.addWidget(self.btn_ocr)

        self.setLayout(layout)

    def reload_devices(self):
        try:
            devices = self.scanner.list_devices()
            self.device_combo.clear()
            for d in devices:
                self.device_combo.addItem(d.Properties["Name"].Value)
            if devices:
                self.scanner.connect_device(0)
        except Exception as e:
            QMessageBox.critical(self, "Chyba", f"Nenašel jsem žádný skener:\n{e}")

    def scan_pages(self):
        self.scanned_images.clear()
        while True:
            if not self.scan_page_single():
                break
            if QMessageBox.question(self, "Další", "Další stránka?") != QMessageBox.StandardButton.Yes:
                break
        if self.scanned_images:
            self.btn_ocr.setEnabled(True)

    def scan_page_single(self):
        index = self.device_combo.currentIndex()
        try:
            self.scanner.connect_device(index)
        except Exception as e:
            QMessageBox.critical(self, "Chyba", f"Nelze připojit ke skeneru:\n{e}")
            return False

        dpi = self.dpi_spin.value()
        source = self.source_combo.currentText()
        color_text = self.color_combo.currentText()
        color_mode = 1 if color_text == "Barevný" else 2 if color_text == "Šedý" else 4

        self.progress_dialog = QProgressDialog("Skenuji...", "Zrušit", 0, 0, self)
        self.progress_dialog.show()

        self.scan_thread = ScanThread(self.scanner, dpi, color_mode, source)
        self.scan_thread.finished.connect(self.scan_finished)
        self.scan_thread.start()

        loop = QEventLoop()
        self.scan_thread.finished.connect(loop.quit)
        loop.exec()

        return hasattr(self, 'last_scanned_image')

    def scan_finished(self, img):
        self.progress_dialog.cancel()
        preview = PreviewDialog(img)
        if preview.exec():
            self.scanned_images.append(preview.image)
            self.last_scanned_image = True
        else:
            self.last_scanned_image = False

    def run_ocr(self):
        self.progress_dialog = QProgressDialog("Probíhá OCR (PaddleOCR)...", "Zrušit", 0, 100, self)
        self.progress_dialog.show()

        lang = self.lang_combo.currentText()
        self.ocr_thread = OCRThread(self.scanned_images, lang)
        self.ocr_thread.finished.connect(self.ocr_finished)
        self.ocr_thread.ocr_results.connect(self.set_ocr_results)
        self.ocr_thread.start()

    def set_ocr_results(self, results):
        self.last_ocr_results = results

    def ocr_finished(self, full_text):
        self.progress_dialog.cancel()
        path, _ = QFileDialog.getSaveFileName(self, "Uložit", "", "Text (*.txt);;PDF (*.pdf);;Word (*.docx)")
        if not path: return

        if path.lower().endswith(".pdf"):
            self.save_pdf(path)
        elif path.lower().endswith(".docx"):
            self.save_docx(full_text, path)
        else:
            with open(path, "w", encoding="utf-8") as f: f.write(full_text)
            QMessageBox.information(self, "Hotovo", "Uloženo.")

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
