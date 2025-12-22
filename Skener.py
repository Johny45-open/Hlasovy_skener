import sys
import tempfile
import os
import io
import threading
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel,
    QComboBox, QSpinBox, QMessageBox, QFileDialog, QDialog,
    QProgressDialog
)
from PyQt6.QtGui import QPixmap
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QEventLoop
from PIL import Image, ImageQt
import pytesseract
import comtypes.client
from docx import Document
import fitz  # PyMuPDF
from gtts import gTTS
import pygame

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

# ------------------ WIA skener ------------------
class WIAScanner:
    def __init__(self):
        self.dm = comtypes.client.CreateObject("WIA.DeviceManager")
        self.device = None

    def list_devices(self):
        return [self.dm.DeviceInfos.Item(i + 1) for i in range(self.dm.DeviceInfos.Count)]

    def connect_device(self, index=0):
        devices = self.list_devices()
        if not devices:
            raise RuntimeError("Žádný skener nenalezen")
        if index < 0 or index >= len(devices):
            raise RuntimeError("Neplatný index skeneru")
        self.device = devices[index].Connect()

    def scan(self, dpi=300, color_mode=1, source="Sklo"):
        if not self.device:
            raise RuntimeError("Skener není připojen")

        def set_prop(name, val):
            for p in self.device.Properties:
                if p.Name == name:
                    p.Value = val
                    return True
            return False

        set_prop("6147", dpi)
        set_prop("6148", dpi)
        source_val = 1 if source == "Sklo" else 2
        set_prop("6146", source_val)
        set_prop("6151", color_mode)

        item = self.device.Items[1]
        image = item.Transfer()

        with tempfile.NamedTemporaryFile(suffix=".bmp") as tmp:
            image.SaveFile(tmp.name)
            with Image.open(tmp.name) as img:
                pil_img = img.copy()
        return pil_img

# ------------------ Náhled dialog ------------------
class PreviewDialog(QDialog):
    def __init__(self, pil_img):
        super().__init__()
        self.setWindowTitle("Náhled skenu")
        self.image = pil_img

        self.label = QLabel()
        self.label.setAccessibleName("Náhled naskenované stránky")
        self.label.setAccessibleDescription("Zobrazení naskenované stránky, lze otočit")
        self.update_image()

        btn_rotate = QPushButton("Otočit")
        btn_rotate.setAccessibleName("Otočit")
        btn_rotate.setAccessibleDescription("Otočí obrázek o 90 stupňů")
        btn_rotate.clicked.connect(self.rotate_image)

        btn_ok = QPushButton("Potvrdit náhled")
        btn_ok.setAccessibleName("Potvrdit náhled")
        btn_ok.setAccessibleDescription("Potvrdí a uloží upravený obrázek")
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

# ------------------ Vlákno skenování ------------------
class ScanThread(QThread):
    finished = pyqtSignal(object)  # PIL obrázek
    error = pyqtSignal(str)

    def __init__(self, scanner, dpi, color_mode, source):
        super().__init__()
        self.scanner = scanner
        self.dpi = dpi
        self.color_mode = color_mode
        self.source = source

    def run(self):
        try:
            img = self.scanner.scan(dpi=self.dpi, color_mode=self.color_mode, source=self.source)
            self.finished.emit(img)
        except Exception as e:
            self.error.emit(str(e))

# ------------------ Vlákno OCR ------------------
class OCRThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)

    def __init__(self, images, lang_code):
        super().__init__()
        self.images = images
        self.lang_code = lang_code

    def run(self):
        full_text = ""
        total = len(self.images)
        for i, img in enumerate(self.images):
            text = pytesseract.image_to_string(img, lang=self.lang_code)
            full_text += f"--- Stránka {i+1} ---\n{text}\n\n"
            self.progress.emit(int((i + 1) / total * 100))
        self.finished.emit(full_text)

# ------------------ Hlavní aplikace ------------------
class ScanApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("WIA skener + OCR + hlas")
        self.scanner = WIAScanner()
        self.scanned_images = []

        layout = QVBoxLayout()

        self.device_combo = QComboBox()
        self.device_combo.setAccessibleName("Výběr skeneru")
        self.device_combo.setAccessibleDescription("Vyber svůj připojený skener")
        self.reload_devices()

        self.source_combo = QComboBox()
        self.source_combo.addItems(["Sklo", "Podavač"])
        self.source_combo.setAccessibleName("Zdroj papíru")
        self.source_combo.setAccessibleDescription("Vyber, jestli skenuješ ze skla nebo podavače")

        self.dpi_spin = QSpinBox()
        self.dpi_spin.setRange(100, 1200)
        self.dpi_spin.setValue(300)
        self.dpi_spin.setSuffix(" DPI")
        self.dpi_spin.setAccessibleName("Rozlišení DPI")
        self.dpi_spin.setAccessibleDescription("Nastav rozlišení skenování v DPI")

        self.color_combo = QComboBox()
        self.color_combo.addItems(["Barevný", "Šedý", "ČB"])
        self.color_combo.setAccessibleName("Režim barev")
        self.color_combo.setAccessibleDescription("Vyber barevný, šedý nebo černobílý režim skenování")

        self.lang_combo = QComboBox()
        self.lang_combo.addItems([
            "ces (Čeština)", "eng (Angličtina)", "deu (Němčina)",
            "fra (Francouzština)", "ita (Italština)", "pol (Polština)"
        ])
        self.lang_combo.setAccessibleName("Jazyk OCR")
        self.lang_combo.setAccessibleDescription("Vyber jazyk, kterým je dokument napsán")

        layout.addWidget(QLabel("Jazyk OCR:"))
        layout.addWidget(self.lang_combo)

        self.btn_scan = QPushButton("Skenovat stránku")
        self.btn_scan.setAccessibleName("Skenovat")
        self.btn_scan.setAccessibleDescription("Spustí skenování stránky")
        self.btn_scan.clicked.connect(self.scan_pages)

        self.btn_ocr = QPushButton("Spustit OCR a uložit")
        self.btn_ocr.setAccessibleName("Spustit OCR")
        self.btn_ocr.setAccessibleDescription("Spustí OCR nad naskenovanými stránkami a umožní uložit text")
        self.btn_ocr.clicked.connect(self.run_ocr)
        self.btn_ocr.setEnabled(False)

        layout.addWidget(QLabel("Vyber skener:"))
        layout.addWidget(self.device_combo)
        layout.addWidget(QLabel("Zdroj papíru:"))
        layout.addWidget(self.source_combo)
        layout.addWidget(QLabel("Rozlišení DPI:"))
        layout.addWidget(self.dpi_spin)
        layout.addWidget(QLabel("Režim:"))
        layout.addWidget(self.color_combo)
        layout.addWidget(self.btn_scan)
        layout.addWidget(self.btn_ocr)

        self.setLayout(layout)

        self.ocr_thread = None
        self.progress_dialog = None

    # ------------------ Skenery ------------------
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

    # ------------------ Skenování ------------------
    def scan_pages(self):
        self.scanned_images.clear()
        while True:
            if not self.scan_page_single():
                break
            odpoved = QMessageBox.question(
                self, "Další stránka",
                "Chceš skenovat další stránku?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if odpoved != QMessageBox.StandardButton.Yes:
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

        self.progress_dialog = QProgressDialog("Probíhá skenování...", "Zrušit", 0, 0, self)
        self.progress_dialog.setWindowModality(Qt.WindowModality.ApplicationModal)
        self.progress_dialog.setMinimumDuration(0)
        self.progress_dialog.canceled.connect(self.cancel_scan)
        self.progress_dialog.show()

        # Spustíme vlákno
        self.scan_thread = ScanThread(self.scanner, dpi, color_mode, source)
        self.scan_thread.finished.connect(self.scan_finished)
        self.scan_thread.error.connect(self.scan_error)
        self.scan_thread.start()

        # Modalní loop dokud neskončí sken
        loop = QEventLoop()
        self.scan_thread.finished.connect(loop.quit)
        self.scan_thread.error.connect(loop.quit)
        loop.exec()

        return hasattr(self, 'last_scanned_image')

    def scan_finished(self, img):
        self.progress_dialog.cancel()
        preview = PreviewDialog(img)
        if preview.exec():
            self.scanned_images.append(preview.image)
            self.last_scanned_image = True
            speak("Stránka naskenována a upravena")
            QMessageBox.information(self, "Hotovo", "Stránka naskenována a upravena.")
        else:
            self.last_scanned_image = False

    def scan_error(self, msg):
        self.progress_dialog.cancel()
        QMessageBox.critical(self, "Chyba", f"Chyba při skenování:\n{msg}")
        self.last_scanned_image = False

    def cancel_scan(self):
        if hasattr(self, 'scan_thread') and self.scan_thread.isRunning():
            self.scan_thread.terminate()
            QMessageBox.information(self, "Zrušeno", "Skenování bylo zrušeno uživatelem.")

    # ------------------ OCR ------------------
    def run_ocr(self):
        if not self.scanned_images:
            QMessageBox.information(self, "Info", "Nejsou žádné naskenované stránky.")
            return

        self.progress_dialog = QProgressDialog("Probíhá OCR...", "Zrušit", 0, 100, self)
        self.progress_dialog.setWindowModality(Qt.WindowModality.ApplicationModal)
        self.progress_dialog.setMinimumDuration(0)
        self.progress_dialog.canceled.connect(self.ocr_cancel)
        self.progress_dialog.show()

        selected_lang = self.lang_combo.currentText().split()[0]
        self.ocr_thread = OCRThread(self.scanned_images, selected_lang)
        self.ocr_thread.finished.connect(self.ocr_finished)
        self.ocr_thread.start()

    def ocr_cancel(self):
        if self.ocr_thread and self.ocr_thread.isRunning():
            self.ocr_thread.terminate()
            QMessageBox.information(self, "Zrušeno", "OCR bylo zrušeno uživatelem.")

    def ocr_finished(self, full_text):
        self.progress_dialog.cancel()
        path, _ = QFileDialog.getSaveFileName(self, "Uložit OCR výstup", "", "Text (*.txt);;PDF (*.pdf);;Word (*.docx)")
        if not path:
            return

        try:
            if path.lower().endswith(".pdf"):
                self.save_pdf(full_text, path)
            elif path.lower().endswith(".docx"):
                self.save_docx(full_text, path)
            else:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(full_text)
                speak("OCR uloženo do TXT")
                QMessageBox.information(self, "Hotovo", "OCR uložen do TXT.")
        except Exception as e:
            QMessageBox.critical(self, "Chyba", f"Chyba při ukládání:\n{e}")

    # ------------------ Uložení PDF ------------------
    def save_pdf(self, full_text, path):
        try:
            dpi = self.dpi_spin.value()
            doc = fitz.open()
            max_width = 1200

            for i, img in enumerate(self.scanned_images):
                if img.width > max_width:
                    ratio = max_width / img.width
                    new_height = int(img.height * ratio)
                    img_resized = img.resize((max_width, new_height), Image.LANCZOS)
                    dpi_adj = dpi * ratio
                else:
                    img_resized = img
                    dpi_adj = dpi

                width_pt = img_resized.width / dpi_adj * 72
                height_pt = img_resized.height / dpi_adj * 72
                page = doc.new_page(width=width_pt, height=height_pt)

                img_byte_arr = io.BytesIO()
                img_resized.save(img_byte_arr, format="JPEG", quality=100)
                size_kb = len(img_byte_arr.getvalue()) / 1024

                jpeg_quality = 85
                if size_kb > 500: jpeg_quality = 50
                elif size_kb > 300: jpeg_quality = 60
                elif size_kb > 200: jpeg_quality = 70

                with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as temp_jpg:
                    img_resized.save(temp_jpg.name, format="JPEG", quality=jpeg_quality)
                    temp_jpg_path = temp_jpg.name

                rect = fitz.Rect(0, 0, width_pt, height_pt)
                page.insert_image(rect, filename=temp_jpg_path)

                words = pytesseract.image_to_data(
                    img_resized,
                    lang=self.lang_combo.currentText().split()[0],
                    output_type=pytesseract.Output.DICT
                )
                for j in range(len(words['text'])):
                    word = words['text'][j]
                    if not word.strip():
                        continue
                    x, y, w, h = words['left'][j], words['top'][j], words['width'][j], words['height'][j]
                    x_pt = x / dpi_adj * 72
                    y_pt = y / dpi_adj * 72
                    w_pt = w / dpi_adj * 72
                    h_pt = h / dpi_adj * 72
                    bbox = fitz.Rect(x_pt, y_pt, x_pt + w_pt, y_pt + h_pt)
                    page.insert_textbox(bbox, word, fontsize=h_pt * 0.8, render_mode=3, fill_opacity=0)

                os.unlink(temp_jpg_path)

            doc.save(path)
            speak("OCR + obrázek uložen do PDF")
            QMessageBox.information(self, "Hotovo", "OCR + obrázek uložen do prohledávatelného PDF.")
        except Exception as e:
            QMessageBox.critical(self, "Chyba", f"Nepodařilo se uložit PDF:\n{e}")

    # ------------------ Uložení DOCX ------------------
    def save_docx(self, text, path):
        doc = Document()
        for part in text.split("\n\n"):
            doc.add_paragraph(part)
            doc.add_page_break()
        doc.save(path)
        speak("OCR uložen do DOCX")
        QMessageBox.information(self, "Hotovo", "OCR uložen do DOCX.")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = ScanApp()
    win.resize(400, 500)
    win.show()
    sys.exit(app.exec())
