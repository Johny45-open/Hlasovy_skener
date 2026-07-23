from __future__ import annotations
from dataclasses import dataclass, field
from typing import Optional, TYPE_CHECKING

from PIL import Image
from PyQt6.QtWidgets import QMessageBox, QFileDialog

from ocr_engine import OcrResult

if TYPE_CHECKING:
    from Skener import ScanApp


@dataclass
class PipelineContext:
    images: list[Image.Image] = field(default_factory=list)
    ocr_results: list[list[OcrResult]] = field(default_factory=list)
    text: str = ""
    cancelled: bool = False


class BaseAction:
    id: str = ""
    name: str = ""
    description: str = ""

    def run(self, ctx: PipelineContext, app: ScanApp) -> None:
        raise NotImplementedError


class ScanAllAction(BaseAction):
    id = "scan_all"
    name = "Skenovat všechny stránky"
    description = "Naskenuje všechny stránky z podavače nebo skla. V dávkovém režimu."

    def run(self, ctx: PipelineContext, app: ScanApp) -> None:
        old_batch = app.batch_cb.isChecked()
        app.batch_cb.setChecked(True)
        app.scan_pages()
        app.batch_cb.setChecked(old_batch)
        ctx.images = list(app.scanned_images)


class ScanAction(BaseAction):
    id = "scan"
    name = "Skenovat jednu stránku"
    description = "Naskenuje jednu stránku. Po každé se zeptá, zda přidat další."

    def run(self, ctx: PipelineContext, app: ScanApp) -> None:
        app.scan_pages()
        ctx.images = list(app.scanned_images)


class PreprocessAction(BaseAction):
    id = "preprocess"
    name = "Předzpracovat obraz"
    description = "Doostří a zvýší kontrast naskenovaných obrázků pro lepší OCR."

    def run(self, ctx: PipelineContext, app: ScanApp) -> None:
        from Skener import preprocess_image
        ctx.images = [preprocess_image(img) for img in ctx.images]
        app.scanned_images = list(ctx.images)
        app.speak(f"Předzpracováno {len(ctx.images)} obrázků.")


class OcrAction(BaseAction):
    id = "ocr"
    name = "OCR rozpoznání"
    description = "Rozpozná text z naskenovaných obrázků pomocí zvoleného OCR enginu."

    def run(self, ctx: PipelineContext, app: ScanApp) -> None:
        if not ctx.images:
            app.speak("Nejsou žádné obrázky k rozpoznání.")
            return
        app._execute_ocr()
        ctx.ocr_results = list(app.last_ocr_results)
        ctx.text = app._last_text


class SaveTxtAction(BaseAction):
    id = "save_txt"
    name = "Uložit jako TXT"
    description = "Zeptá se na umístění a uloží rozpoznaný text jako textový soubor."

    def run(self, ctx: PipelineContext, app: ScanApp) -> None:
        if not ctx.text:
            app.speak("Není žádný text k uložení.")
            return
        path, _ = QFileDialog.getSaveFileName(
            app, "Uložit jako TXT", "", "Text (*.txt)"
        )
        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write(ctx.text)
            app.speak("Text uložen.")


class SavePdfAction(BaseAction):
    id = "save_pdf"
    name = "Uložit jako PDF"
    description = "Zeptá se na umístění a uloží obrázky s vyhledávatelným textem do PDF."

    def run(self, ctx: PipelineContext, app: ScanApp) -> None:
        if not ctx.images:
            app.speak("Nejsou žádné obrázky k uložení.")
            return
        path, _ = QFileDialog.getSaveFileName(
            app, "Uložit jako PDF", "", "PDF (*.pdf)"
        )
        if path:
            app._save_pdf(path)
            app.speak("PDF uloženo.")


class SaveDocxAction(BaseAction):
    id = "save_docx"
    name = "Uložit jako DOCX"
    description = "Zeptá se na umístění a uloží rozpoznaný text jako Word dokument."

    def run(self, ctx: PipelineContext, app: ScanApp) -> None:
        if not ctx.text:
            app.speak("Není žádný text k uložení.")
            return
        path, _ = QFileDialog.getSaveFileName(
            app, "Uložit jako DOCX", "", "Word (*.docx)"
        )
        if path:
            app._save_docx(ctx.text, path)
            app.speak("DOCX uloženo.")


class ExportImagesAction(BaseAction):
    id = "export_images"
    name = "Exportovat obrázky"
    description = "Zeptá se na formát a složku a uloží naskenované obrázky."

    def run(self, ctx: PipelineContext, app: ScanApp) -> None:
        app.export_images()


class ReadTextAction(BaseAction):
    id = "read_text"
    name = "Přečíst text hlasem"
    description = "Přečte rozpoznaný text pomocí screen readeru."

    def run(self, ctx: PipelineContext, app: ScanApp) -> None:
        if ctx.text:
            app.speak(ctx.text)
        else:
            app.speak("Není žádný text k přečtení.")


class ClearPagesAction(BaseAction):
    id = "clear_pages"
    name = "Smazat stránky"
    description = "Smaže všechny naskenované stránky a OCR výsledky."

    def run(self, ctx: PipelineContext, app: ScanApp) -> None:
        app.clear_pages()
        ctx.images.clear()
        ctx.ocr_results.clear()
        ctx.text = ""
        app.speak("Všechny stránky smazány.")


class AskContinueAction(BaseAction):
    id = "ask_continue"
    name = "Zeptat se na pokračování"
    description = "Zobrazí dialog s otázkou, zda pokračovat v provádění makra."

    def run(self, ctx: PipelineContext, app: ScanApp) -> None:
        msg = QMessageBox(app)
        msg.setWindowTitle("Pokračovat?")
        msg.setText("Pokračovat v provádění makra?")
        btn_yes = msg.addButton("Ano", QMessageBox.ButtonRole.YesRole)
        btn_yes.setAccessibleName("Ano, pokračovat v makru")
        btn_no = msg.addButton("Ne", QMessageBox.ButtonRole.NoRole)
        btn_no.setAccessibleName("Ne, zastavit makro")
        msg.exec()
        ctx.cancelled = msg.clickedButton() != btn_yes


# ---------- Registry ----------

_actions: dict[str, BaseAction] = {}


def register(action: BaseAction) -> None:
    _actions[action.id] = action


def get_action(action_id: str) -> Optional[BaseAction]:
    return _actions.get(action_id)


def list_actions() -> list[BaseAction]:
    return list(_actions.values())


# Register built-in actions
register(ScanAllAction())
register(ScanAction())
register(PreprocessAction())
register(OcrAction())
register(SaveTxtAction())
register(SavePdfAction())
register(SaveDocxAction())
register(ExportImagesAction())
register(ReadTextAction())
register(ClearPagesAction())
register(AskContinueAction())
