import threading
import os
from dataclasses import dataclass
from typing import Optional

import numpy as np
import pytesseract
import easyocr
from PIL import Image
from PyQt6.QtCore import QThread, pyqtSignal


@dataclass
class OcrResult:
    text: str
    bbox: Optional[list[list[float]]] = None


LANG_MAP_EASYOCR = {
    "ces": "cs", "eng": "en", "deu": "de",
    "fra": "fr", "ita": "it", "pol": "pl",
}


def _extract_lang_code(lang_combo_text: str) -> str:
    return lang_combo_text.split()[0]


class TesseractThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)
    ocr_results = pyqtSignal(list)

    def __init__(self, images: list[Image.Image], lang_code: str) -> None:
        super().__init__()
        self.images = images
        self.lang_code = _extract_lang_code(lang_code)

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
            full_text += f"--- Stránka {i + 1} ---\n{page_text}\n\n"
            results_list.append(page_results)

            self.progress.emit(int((i + 1) / total * 100))

        self.ocr_results.emit(results_list)
        self.finished.emit(full_text)


class EasyOCRThread(QThread):
    _lock = threading.Lock()
    _reader = None
    _reader_lang = None

    progress = pyqtSignal(int)
    finished = pyqtSignal(str)
    ocr_results = pyqtSignal(list)
    model_loading = pyqtSignal(int)

    def __init__(self, images: list[Image.Image], lang_code: str) -> None:
        super().__init__()
        self.images = images
        pyt_code = _extract_lang_code(lang_code)
        self.lang_code = LANG_MAP_EASYOCR.get(pyt_code, pyt_code[:2])

    def run(self) -> None:
        self._ensure_reader()

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

            full_text += f"--- Stránka {i + 1} ---\n{page_text}\n\n"
            results_list.append(page_results)

            self.progress.emit(int((i + 1) / total * 100))

        self.ocr_results.emit(results_list)
        self.finished.emit(full_text)

    def _ensure_reader(self) -> None:
        with EasyOCRThread._lock:
            if EasyOCRThread._reader is None or \
               EasyOCRThread._reader_lang != [self.lang_code]:
                self.model_loading.emit(0)
                EasyOCRThread._reader = easyocr.Reader(
                    [self.lang_code],
                    gpu=False,
                    model_storage_directory=os.path.join(
                        os.path.dirname(os.path.abspath(__file__)),
                        ".easyocr_model",
                    ),
                )
                EasyOCRThread._reader_lang = [self.lang_code]
                self.model_loading.emit(100)


def create_ocr_thread(
    engine: str, images: list[Image.Image], lang_code: str
) -> TesseractThread | EasyOCRThread:
    if engine == "EasyOCR":
        return EasyOCRThread(images, lang_code)
    return TesseractThread(images, lang_code)
