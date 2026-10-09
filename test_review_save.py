"""Testy kontroly před exportem a sjednoceného zdroje (kroky 4+5).

GUI testy běží offscreen (QT_QPA_PLATFORM=offscreen), dialogy jsou mockované.
- Accept → potvrzení + uložení; reject → žádný export, pending vyčištěn.
- Retry → nový pokus, pak znovu kontrola (ne automatické uložení).
- Neplatná struktura → oprava / zrušení, nikdy tichý přesun ani ztráta.
- Platné potvrzení → dialog se znovu neotevírá.
- TXT/DOCX/PDF respektují stejnou potvrzenou verzi.
"""
import os

import pytest
from PIL import Image
from PyQt6.QtWidgets import QApplication, QDialog

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import Skener
from Skener import ScanApp, OcrPreviewDialog
from ocr_engine import OcrResult


@pytest.fixture(scope="module")
def qapp():
    app = QApplication.instance() or QApplication([])
    yield app


class FakeTextEdit:
    def setAccessibleName(self, *a):
        pass

    def setAccessibleDescription(self, *a):
        pass


class FakePreviewDialog:
    """Fronta (result, text) – mock Skener.OcrPreviewDialog."""
    RESULT_RETRY = OcrPreviewDialog.RESULT_RETRY
    script = []
    opened = 0
    last_title = ""

    def __init__(self, text, parent=None, alt_engine_name=""):
        self._text = text
        self.text_edit = FakeTextEdit()
        FakePreviewDialog.opened += 1

    def setWindowTitle(self, title):
        FakePreviewDialog.last_title = title

    def exec(self):
        res, text = FakePreviewDialog.script.pop(0)
        self._text = text
        return res

    def get_text(self):
        return self._text


@pytest.fixture()
def app(qapp, monkeypatch):
    monkeypatch.setattr(Skener, "speak", lambda *a, **k: None)
    monkeypatch.setattr(Skener, "OcrPreviewDialog", FakePreviewDialog)
    FakePreviewDialog.script = []
    FakePreviewDialog.opened = 0
    FakePreviewDialog.last_title = ""
    w = ScanApp()
    # Realistická velikost skenu při 300 DPI (1200x1600 jako ostatní PDF testy).
    img = Image.new("RGB", (1200, 1600), "white")
    w.scanned_images = [img, img]
    w.raw_scanned_images = [img, img]
    w.page_list.addItems(["Stránka 1", "Stránka 2"])

    def bbox(x=50, y=50):
        return [[float(x), float(y)], [float(x + 400), float(y)],
                [float(x + 400), float(y + 60)], [float(x), float(y + 60)]]

    w.last_ocr_results = [
        [OcrResult(text="Alfa", bbox=bbox(), original_text="Alfa", processed_text="Alfa")],
        [OcrResult(text="Beta", bbox=bbox(), original_text="Beta", processed_text="Beta")],
    ]
    w._last_text = "--- Stránka 1 ---\nAlfa\n\n--- Stránka 2 ---\nBeta\n\n"
    w._last_edited_pages = ["Alfa", "Beta"]
    w._ocr_rev = 1
    w._invalidate_confirmation()
    saved = []
    monkeypatch.setattr(w, "_do_save_ocr_format", lambda fmt: saved.append(fmt))
    monkeypatch.setattr(w, "_show_save_success", lambda *a, **k: None)
    w._saved_formats = saved
    return w


def _accept(text):
    return (QDialog.DialogCode.Accepted, text)


def _reject(text=""):
    return (QDialog.DialogCode.Rejected, text)


# --- krok 4: kontrola před uložením ---

def test_accept_confirms_and_saves(app):
    FakePreviewDialog.script = [_accept(app._last_text)]
    app._pending_save_format = "pdf_ocr"
    app._try_save_with_ocr("pdf_ocr")
    assert app._saved_formats == ["pdf_ocr"]
    assert app.is_confirmation_valid()
    assert app._confirmed_pages == ["Alfa", "Beta"]


def test_reject_cancels_without_export(app):
    FakePreviewDialog.script = [_reject()]
    app._pending_save_format = "pdf_ocr"
    app._try_save_with_ocr("pdf_ocr")
    assert app._saved_formats == []
    assert app._pending_save_format is None
    assert not app.is_confirmation_valid()


def test_retry_runs_new_ocr_then_reviews_again(app, monkeypatch):
    calls = []

    def fake_retry(interactive=True):
        calls.append(interactive)
        app._last_text = "--- Stránka 1 ---\nAlfa2\n\n--- Stránka 2 ---\nBeta2\n\n"

    monkeypatch.setattr(app, "_retry_with_alternate_engine", fake_retry)
    FakePreviewDialog.script = [
        (OcrPreviewDialog.RESULT_RETRY, app._last_text),
        _accept("--- Stránka 1 ---\nAlfa2\n\n--- Stránka 2 ---\nBeta2\n\n"),
    ]
    app._pending_save_format = "pdf_ocr"
    app._try_save_with_ocr("pdf_ocr")
    assert calls == [True]
    assert FakePreviewDialog.opened == 2
    assert app._saved_formats == ["pdf_ocr"]
    assert app._confirmed_pages == ["Alfa2", "Beta2"]


def test_invalid_structure_fix_then_save(app, monkeypatch):
    monkeypatch.setattr(app, "_ask_confirmed_pages_invalid", lambda total: True)
    bad = "souvislý text bez headerů"
    good = "--- Stránka 1 ---\nAlfa\n\n--- Stránka 2 ---\nBeta\n\n"
    FakePreviewDialog.script = [_accept(bad), _accept(good)]
    app._pending_save_format = "txt"
    app._try_save_with_ocr("txt")
    assert app._saved_formats == ["txt"]
    assert app._confirmed_pages == ["Alfa", "Beta"]


def test_invalid_structure_cancel_no_export_no_loss(app, monkeypatch):
    monkeypatch.setattr(app, "_ask_confirmed_pages_invalid", lambda total: False)
    bad = "souvislý text bez headerů"
    FakePreviewDialog.script = [_accept(bad)]
    before = list(app.last_ocr_results)
    app._pending_save_format = "txt"
    app._try_save_with_ocr("txt")
    assert app._saved_formats == []
    assert app._pending_save_format is None
    # Původní OCR výsledek zůstal nedotčen (nic se nezahodilo).
    assert app.last_ocr_results == before
    assert "Alfa" in app._last_text


def test_valid_confirmation_skips_dialog(app):
    app._confirm_pages(["Alfa", "Beta"], app._last_text)
    app._pending_save_format = "docx"
    app._try_save_with_ocr("docx")
    assert FakePreviewDialog.opened == 0
    assert app._saved_formats == ["docx"]


def test_new_ocr_invalidates_confirmation(app):
    app._confirm_pages(["Alfa", "Beta"], app._last_text)
    assert app.is_confirmation_valid()
    app._invalidate_confirmation()
    assert not app.is_confirmation_valid()


def test_confirmation_invalid_after_page_count_change(app):
    app._confirm_pages(["Alfa", "Beta"], app._last_text)
    app.scanned_images.append(Image.new("RGB", (1200, 1600), "white"))
    assert not app.is_confirmation_valid()


def test_confirmation_invalid_after_rev_bump(app):
    app._confirm_pages(["Alfa", "Beta"], app._last_text)
    app._ocr_rev += 1
    assert not app.is_confirmation_valid()


# --- krok 5: stejná potvrzená verze pro TXT/DOCX/PDF ---

def test_txt_docx_pdf_share_confirmed_version(app, tmp_path):
    app._confirm_pages(["Alfa opravená", "Beta opravená"],
                       "--- Stránka 1 ---\nAlfa opravená\n\n--- Stránka 2 ---\nBeta opravená\n\n")
    txt_path = str(tmp_path / "out.txt")
    app._save_txt_pages(txt_path)
    txt = open(txt_path, encoding="utf-8").read()
    assert "--- Stránka 1 ---\nAlfa opravená" in txt
    assert "--- Stránka 2 ---\nBeta opravená" in txt

    docx_path = str(tmp_path / "out.docx")
    app._save_docx_pages(docx_path)
    from docx import Document
    doc = Document(docx_path)
    full = "\n".join(p.text for p in doc.paragraphs)
    assert "Alfa opravená" in full
    assert "Beta opravená" in full
    assert "Alfa\n" not in full.replace("Alfa opravená", "")

    from Skener import build_confirmed_pdf_plan, build_searchable_pdf
    import tempfile
    override, changed = build_confirmed_pdf_plan(
        app.last_ocr_results, app.get_export_pages(), 2)
    assert changed == [True, True]
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    tmp.close()
    try:
        build_searchable_pdf(app.scanned_images, app.last_ocr_results,
                             300, tmp.name, confirmed_override=override)
        import fitz
        doc = fitz.open(tmp.name)
        try:
            t0 = doc[0].get_text()
            t1 = doc[1].get_text()
        finally:
            doc.close()
        assert "Alfa opravená" in " ".join(t0.split())
        assert "Beta opravená" in " ".join(t1.split())
        assert " ".join(t0.split()).count("Alfa opravená") == 1
        # Starý text není duplicitně ponechán.
        assert " ".join(t0.split()).replace("Alfa opravená", "").strip() == ""
    finally:
        os.remove(tmp.name)


def test_get_export_pages_falls_back_without_confirmation(app):
    assert app.get_export_pages() == ["Alfa", "Beta"]


def test_save_pdf_with_ocr_uses_confirmed_text_e2e(app, tmp_path):
    """Auditní repro A1: oprava z náhledu se musí objevit ve vrstvě PDF."""
    app._confirm_pages(["Alfa opravená", "Beta"],
                       "--- Stránka 1 ---\nAlfa opravená\n\n--- Stránka 2 ---\nBeta\n\n")
    out = str(tmp_path / "potvrzeno.pdf")
    app._save_pdf_with_ocr(out)
    import fitz
    doc = fitz.open(out)
    try:
        assert len(doc) == 2
        t0 = " ".join(doc[0].get_text().split())
        t1 = " ".join(doc[1].get_text().split())
        assert "Alfa opravená" in t0
        assert t0.replace("Alfa opravená", "").strip() == ""
        assert "Beta" in t1
        assert len(doc[0].get_images(full=True)) >= 1
    finally:
        doc.close()
