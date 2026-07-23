from PyQt6.QtCore import QThread, pyqtSignal
import pytesseract
import numpy as np
import easyocr
import os

# Mapování kódů jazyků z Tesseract/UI do EasyOCR
LANG_MAP_EASYOCR = {
    "ces": "cs",
    "eng": "en",
    "deu": "de",
    "fra": "fr",
    "ita": "it",
    "pol": "pl",
}


class TesseractThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)
    ocr_results = pyqtSignal(list)

    def __init__(self, images, lang_code):
        super().__init__()
        self.images = images
        self.lang_code = lang_code.split()[0]

    def run(self):
        full_text = ""
        results_list = []
        total = len(self.images)
        for i, img in enumerate(self.images):
            text = pytesseract.image_to_string(img, lang=self.lang_code)

            data = pytesseract.image_to_data(img, lang=self.lang_code, output_type=pytesseract.Output.DICT)

            page_results = []
            n_boxes = len(data['text'])
            for j in range(n_boxes):
                if int(data['conf'][j]) > 0:
                    text_content = data['text'][j]
                    if text_content.strip():
                        bbox = [
                            [data['left'][j], data['top'][j]],
                            [data['left'][j] + data['width'][j], data['top'][j]],
                            [data['left'][j] + data['width'][j], data['top'][j] + data['height'][j]],
                            [data['left'][j], data['top'][j] + data['height'][j]]
                        ]
                        page_results.append({'bbox': bbox, 'text': text_content})

            full_text += f"--- Stránka {i+1} ---\n{text}\n\n"
            results_list.append(page_results)

            self.progress.emit(int((i + 1) / total * 100))

        self.finished.emit(full_text)
        self.ocr_results.emit(results_list)


class EasyOCRThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)
    ocr_results = pyqtSignal(list)
    model_loading = pyqtSignal(int)

    _reader = None
    _reader_lang = None

    def __init__(self, images, lang_code):
        super().__init__()
        self.images = images
        pyt_code = lang_code.split()[0]
        self.lang_code = LANG_MAP_EASYOCR.get(pyt_code, pyt_code[:2])

    def run(self):
        if (EasyOCRThread._reader is None or
                EasyOCRThread._reader_lang != [self.lang_code]):
            self.model_loading.emit(0)
            EasyOCRThread._reader = easyocr.Reader(
                [self.lang_code],
                gpu=False,
                model_storage_directory=os.path.join(
                    os.path.dirname(os.path.abspath(__file__)),
                    ".easyocr_model"
                )
            )
            EasyOCRThread._reader_lang = [self.lang_code]
            self.model_loading.emit(100)

        full_text = ""
        results_list = []
        total = len(self.images)

        for i, img in enumerate(self.images):
            img_np = np.array(img)
            raw_results = EasyOCRThread._reader.readtext(img_np, paragraph=True)

            page_text = ""
            page_results = []
            for text, conf in raw_results:
                if text.strip():
                    page_text += text + "\n"
                    page_results.append({'text': text})

            full_text += f"--- Stránka {i+1} ---\n{page_text}\n\n"
            results_list.append(page_results)

            self.progress.emit(int((i + 1) / total * 100))

        self.finished.emit(full_text)
        self.ocr_results.emit(results_list)


# Zachování zpětné kompatibility
OCRThread = TesseractThread


def create_ocr_thread(engine, images, lang_code):
    """Vrátí příslušný OCR thread podle zvoleného enginu."""
    if engine == "EasyOCR":
        return EasyOCRThread(images, lang_code)
    return TesseractThread(images, lang_code)
