from PyQt6.QtCore import QThread, pyqtSignal
from paddleocr import PaddleOCR
import numpy as np

class OCRThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)
    # Pro PDF generování předáváme i surová data o boxech
    ocr_results = pyqtSignal(list)

    def __init__(self, images, lang_code):
        super().__init__()
        self.images = images
        # PaddleOCR používá kódy jako 'cs', 'en' atd.
        self.lang_code = lang_code.split()[0].replace("ces", "cs").replace("eng", "en")
        self.ocr = PaddleOCR(use_angle_cls=True, lang=self.lang_code)

    def run(self):
        full_text = ""
        results_list = []
        total = len(self.images)
        for i, img in enumerate(self.images):
            # PaddleOCR očekává obrázek jako numpy array (z PIL)
            img_np = np.array(img)
            result = self.ocr.ocr(img_np, cls=True)
            
            page_text = ""
            page_results = []
            
            # PaddleOCR result je seznam detekcí pro jednu stránku
            for line in result[0]:
                bbox = line[0]  # [[x1, y1], [x2, y2], ...]
                text, score = line[1]
                page_text += text + "\n"
                page_results.append({'bbox': bbox, 'text': text})
            
            full_text += f"--- Stránka {i+1} ---\n{page_text}\n\n"
            results_list.append(page_results)
            
            self.progress.emit(int((i + 1) / total * 100))
        
        self.finished.emit(full_text)
        self.ocr_results.emit(results_list)
