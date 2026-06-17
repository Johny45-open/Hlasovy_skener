from PyQt6.QtCore import QThread, pyqtSignal
import pytesseract
import numpy as np

class OCRThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)
    # Pro PDF generování předáváme i surová data o boxech
    ocr_results = pyqtSignal(list)

    def __init__(self, images, lang_code):
        super().__init__()
        self.images = images
        # pytesseract používá ISO 639-2 kódy jako 'ces', 'eng' atd.
        self.lang_code = lang_code.split()[0]

    def run(self):
        full_text = ""
        results_list = []
        total = len(self.images)
        for i, img in enumerate(self.images):
            # Pytesseract očekává obrázek jako PIL Image
            text = pytesseract.image_to_string(img, lang=self.lang_code)
            
            # Pro PDF generování potřebujeme boxy, tedy image_to_data
            data = pytesseract.image_to_data(img, lang=self.lang_code, output_type=pytesseract.Output.DICT)
            
            page_results = []
            # Zpracování data pro extrakci bboxů a textu
            n_boxes = len(data['text'])
            for j in range(n_boxes):
                if int(data['conf'][j]) > 0: # Pokud je nějaký text detekován
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
