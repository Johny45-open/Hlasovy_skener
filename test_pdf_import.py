"""Testy importu existujicich PDF se selektivnim OCR (pdf_import.py).

Strategie: zadny realny OCR engine (Tesseract/EasyOCR se mockuji),
PDF fixtures se generuji programove pres PyMuPDF/PIL. Smycka
render -> OCR -> overlay -> validace se testuje s canned OcrResult;
engine vetve sync OCR se testuji s mockem pytesseract / Readeru.
"""
import io
import os

import fitz
import pytest
from PIL import Image, ImageDraw

import pdf_import
from pdf_import import (
    analyze_pdf,
    build_jobs,
    default_output_path,
    format_analysis_summary,
    overlay_ocr_layer,
    process_single_job,
    validate_imported_pdf,
    ImportJob,
    PdfImportWorker,
    UnsupportedRotationError,
    PDF_IMPORT_RENDER_DPI,
)
from ocr_engine import OcrResult


CZ = "Příliš žluťoučký kůň úpěl ďábelské ódy"
LONG_TEXT = ("Toto je ukázkový text pro testování analýzy PDF. " * 6).strip()
CZ_LONG = (CZ + ". " * 1 + " ") * 4
CZ_LONG = (CZ + " ") * 8  # > 100 znaku vcetne diakritiky

FONTFILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "assets", "fonts", "DejaVuSans.ttf")


def _word_bbox(x, y, w, h):
    return [[float(x), float(y)], [float(x + w), float(y)],
            [float(x + w), float(y + h)], [float(x), float(y + h)]]


def _add_text_page(doc, text):
    page = doc.new_page(width=595, height=842)
    page.insert_font(fontname="t-dejavu", fontfile=FONTFILE)
    # Vloz po radcich, aby se texty neroztekly mimo stranu.
    y = 80.0
    words, line = text.split(), ""
    for w in words:
        if len(line) + len(w) + 1 > 70:
            page.insert_text((72, y), line, fontname="t-dejavu", fontsize=12)
            y += 18
            line = w
        else:
            line = (line + " " + w).strip()
    if line:
        page.insert_text((72, y), line, fontname="t-dejavu", fontsize=12)
    return page


def _add_image_page(doc, label="sken"):
    img = Image.new("RGB", (1200, 1600), "white")
    d = ImageDraw.Draw(img)
    d.text((100, 100), label, fill="black")
    buf = io.BytesIO()
    img.save(buf, "PNG")
    page = doc.new_page(width=595, height=842)
    page.insert_image(page.rect, stream=buf.getvalue())
    return page


def _save_doc(doc, path):
    doc.save(path)
    doc.close()
    return path


def _make_text_pdf(path, pages=3):
    doc = fitz.open()
    for _ in range(pages):
        _add_text_page(doc, LONG_TEXT)
    return _save_doc(doc, path)


def _make_image_pdf(path, pages=2):
    doc = fitz.open()
    for i in range(pages):
        _add_image_page(doc, f"sken {i + 1}")
    return _save_doc(doc, path)


def _make_mixed_pdf(path):
    """8 stran: textove 3, 5, 6 (1-based), ostatni obrazove."""
    doc = fitz.open()
    for i in range(1, 9):
        if i in (3, 5, 6):
            _add_text_page(doc, LONG_TEXT)
        else:
            _add_image_page(doc, f"sken {i}")
    return _save_doc(doc, path)


def _mock_ocr(monkeypatch, text=CZ, pages_map=None):
    """Nah radi ocr_images_sync canned vysledky (bbox v render pixelech)."""
    def fake(images, engine, lang_raw, diacritics_enabled=False,
             progress_cb=None, is_cancelled=None):
        out = []
        for i, _img in enumerate(images):
            t = pages_map[i] if pages_map is not None else text
            out.append([OcrResult(text=t, bbox=_word_bbox(100, 100, 800, 80))])
            if progress_cb is not None:
                progress_cb(i + 1, len(images))
        return out
    monkeypatch.setattr(pdf_import, "ocr_images_sync", fake)
    return fake


def _page_texts(path):
    doc = fitz.open(path)
    try:
        return [p.get_text().strip() for p in doc], len(doc)
    finally:
        doc.close()


# 1 - textove PDF: OCR se nenabidne, nevznikne vystup
def test_text_pdf_needs_no_ocr(tmp_path):
    src = str(tmp_path / "text.pdf")
    _make_text_pdf(src)
    analysis = analyze_pdf(src)
    assert analysis.ok
    assert analysis.page_count == 3
    assert analysis.pages_needing_ocr == []
    for p in analysis.pages:
        assert not p.needs_ocr
        assert p.char_count >= 100
    assert "OCR není potřeba" in format_analysis_summary(analysis)
    jobs, skipped = build_jobs([analysis], "Tesseract", "ces (Čeština)", False)
    assert jobs == []
    assert skipped == []
    assert not os.path.exists(default_output_path(src))


# 2 - image-only PDF: vsechny strany NEEDS_OCR
def test_image_pdf_all_pages_need_ocr(tmp_path):
    src = str(tmp_path / "sken.pdf")
    _make_image_pdf(src, pages=2)
    analysis = analyze_pdf(src)
    assert analysis.ok
    assert analysis.pages_needing_ocr == [1, 2]


# 3 - mixed PDF: OCR jen chybejici strany
def test_mixed_pdf_selective_pages(tmp_path):
    src = str(tmp_path / "faktura.pdf")
    _make_mixed_pdf(src)
    analysis = analyze_pdf(src)
    assert analysis.ok
    assert analysis.page_count == 8
    assert analysis.pages_needing_ocr == [1, 2, 4, 7, 8]
    summary = format_analysis_summary(analysis)
    assert "8 stran" in summary and "5 stranách" in summary


# kratka textova strana: konzervativne NEEDS_OCR (bezpecny smer)
def test_short_text_page_is_conservatively_ocr(tmp_path):
    doc = fitz.open()
    _add_text_page(doc, "Faktura č. 123")
    path = str(tmp_path / "kratke.pdf")
    _save_doc(doc, path)
    analysis = analyze_pdf(path)
    assert analysis.pages_needing_ocr == [1]


# 4 - vice PDF: spravny pocet vystupu, zadny neprepise zdroj
def test_multi_pdf_job_planning(tmp_path):
    srcs = []
    for name in ("a.pdf", "b.pdf"):
        src = str(tmp_path / name)
        _make_image_pdf(src, pages=1)
        srcs.append(src)
    analyses = [analyze_pdf(s) for s in srcs]
    jobs, skipped = build_jobs(analyses, "Tesseract", "ces (Čeština)", False)
    assert len(jobs) == 2
    assert skipped == []
    outs = [j.out_path for j in jobs]
    assert len(set(outs)) == 2
    for j in jobs:
        assert os.path.abspath(j.out_path) != os.path.abspath(j.src_path)
        assert j.out_path.endswith("_OCR.pdf")


# 5/12/13 - diakritika + zachovani textovych stran + obsah OCR stran
def test_mixed_pdf_overlay_preserves_and_adds(tmp_path, monkeypatch):
    src = str(tmp_path / "faktura.pdf")
    _make_mixed_pdf(src)
    before, _ = _page_texts(src)
    _mock_ocr(monkeypatch, text=CZ)
    out = default_output_path(src)
    job = ImportJob(src_path=src, out_path=out, pages_to_ocr=[1, 2, 4, 7, 8],
                    engine="Tesseract", lang_raw="ces (Čeština)")
    result = process_single_job(job)
    assert result.status == "ok", result.message
    assert os.path.isfile(out)
    # Zdroj zmenen nebyl (analyza zdroje stale hlasi potrebu OCR).
    assert analyze_pdf(src).pages_needing_ocr == [1, 2, 4, 7, 8]
    after, n = _page_texts(out)
    assert n == 8
    # Puvodni textove strany obsahove shodne.
    for i in (2, 4, 5):  # 0-based indexy stran 3, 5, 6
        assert after[i] == before[i]
        assert "ukázkový text" in after[i]
    # OCR strany obsahuji ocekavany text vcetne diakritiky, bez '?'.
    for i in (0, 1, 3, 6, 7):
        assert CZ in after[i]
    for ch in ["ř", "š", "ů", "ň", "ď", "ť", "ž", "ý", "á", "í", "é", "ú", "ó"]:
        assert ch in after[0]
    assert "?" not in after[0]


# 13b - nezavisly cross-check pres pdfplumber
def test_pdfplumber_crosscheck(tmp_path, monkeypatch):
    pl = pytest.importorskip("pdfplumber")
    src = str(tmp_path / "sken.pdf")
    _make_image_pdf(src, pages=1)
    _mock_ocr(monkeypatch, text=CZ)
    out = default_output_path(src)
    result = process_single_job(ImportJob(src_path=src, out_path=out, pages_to_ocr=[1]))
    assert result.status == "ok", result.message
    import pdfplumber as _pl
    with _pl.open(out) as pdf:
        assert len(pdf.pages) == 1
        t = pdf.pages[0].extract_text() or ""
    for word in ["Příliš", "žluťoučký", "kůň"]:
        assert word in t


# 6 - Tesseract sync vetev (mock pytesseract)
def test_tesseract_sync_branch(monkeypatch):
    import pytesseract
    data = {"text": ["Hello", ""], "conf": ["90", "-1"],
            "left": [10, 0], "top": [20, 0], "width": [100, 0], "height": [30, 0]}
    monkeypatch.setattr(pytesseract, "image_to_data", lambda *a, **k: data)
    monkeypatch.setattr(pytesseract, "image_to_string", lambda *a, **k: "Hello")
    img = Image.new("RGB", (400, 200), "white")
    pages = pdf_import.tesseract_ocr_images_sync([img], "ces (Čeština)")
    assert len(pages) == 1
    assert pages[0][0].text == "Hello"
    assert pages[0][0].bbox[0][0] == pytest.approx(10.0)


# 7 - EasyOCR sync vetev (mock Reader)
def test_easyocr_sync_branch(monkeypatch):
    class FakeReader:
        def readtext(self, arr, paragraph=True):
            assert paragraph is True
            return [([ [1.0, 2.0], [11.0, 2.0], [11.0, 12.0], [1.0, 12.0] ], "Easy text", 0.9)]
    monkeypatch.setattr(pdf_import, "ensure_easyocr_reader", lambda lang: FakeReader())
    img = Image.new("RGB", (400, 200), "white")
    pages = pdf_import.easyocr_ocr_images_sync([img], "ces (Čeština)")
    assert pages[0][0].text == "Easy text"
    assert pages[0][0].bbox[0][0] == pytest.approx(1.0)
    # mapovani jazyka ces -> cs
    seen = {}
    def fake_ensure(lang):
        seen["lang"] = lang
        return FakeReader()
    monkeypatch.setattr(pdf_import, "ensure_easyocr_reader", fake_ensure)
    pdf_import.easyocr_ocr_images_sync([img], "ces (Čeština)")
    assert seen["lang"] == "cs"


# 8 - poskozene PDF: chyba jednoho souboru neukonci davku
def test_corrupt_pdf_does_not_stop_batch(tmp_path, monkeypatch):
    bad = str(tmp_path / "rozbite.pdf")
    with open(bad, "wb") as f:
        f.write(b"%PDF-toto-neni-pdf\x00\xff\xfe" * 100)
    good = str(tmp_path / "dobre.pdf")
    _make_image_pdf(good, pages=1)
    _mock_ocr(monkeypatch)
    analyses = [analyze_pdf(bad), analyze_pdf(good)]
    assert analyses[0].error  # srozumitelny duvod
    jobs, skipped = build_jobs(analyses, "Tesseract", "ces", False)
    assert len(jobs) == 1 and len(skipped) == 1
    res = process_single_job(ImportJob(src_path=bad, out_path=default_output_path(bad),
                                       pages_to_ocr=[1]))
    assert res.status == "failed"
    res_ok = process_single_job(jobs[0])
    assert res_ok.status == "ok", res_ok.message


# 9 - sifrovane PDF
def test_encrypted_pdf(tmp_path):
    src = str(tmp_path / "tajne.pdf")
    doc = fitz.open()
    _add_text_page(doc, LONG_TEXT)
    doc.save(src, encryption=fitz.PDF_ENCRYPT_AES_256, user_pw="secret", owner_pw="owner")
    doc.close()
    analysis = analyze_pdf(src)
    assert "hesl" in analysis.error.lower()
    res = process_single_job(ImportJob(src_path=src, out_path=default_output_path(src),
                                       pages_to_ocr=[1]))
    assert res.status == "failed"
    assert "hesl" in res.message.lower()


# 10 - existujici vystup + nikdy out == src
def test_existing_output_requires_consent(tmp_path):
    src = str(tmp_path / "dok.pdf")
    _make_image_pdf(src, pages=1)
    analysis = analyze_pdf(src)
    out = default_output_path(src)
    with open(out, "wb") as f:
        f.write(b"dummy")
    jobs, skipped = build_jobs([analysis], "Tesseract", "ces", False, {})
    assert jobs == [] and len(skipped) == 1 and skipped[0].status == "skipped"
    jobs2, _ = build_jobs([analysis], "Tesseract", "ces", False, {out: True})
    assert len(jobs2) == 1
    # out == src je vzdy odmitnuto
    res = process_single_job(ImportJob(src_path=src, out_path=src, pages_to_ocr=[1]))
    assert res.status == "failed"
    assert "nepřepsal" in res.message or "Zdroj" in res.message or "zdroj" in res.message


def test_default_output_naming(tmp_path):
    src = str(tmp_path / "faktura.pdf")
    out = default_output_path(src)
    assert out.endswith("faktura_OCR.pdf")
    assert os.path.dirname(out) == os.path.dirname(os.path.abspath(src))
    # i *_OCR.pdf dostane dalsi suffix, nikdy se nerovna zdroji
    out2 = default_output_path(out)
    assert os.path.abspath(out2) != os.path.abspath(out)


# 11 - zruseni davky
def test_cancel_before_and_during(tmp_path, monkeypatch):
    src = str(tmp_path / "sken.pdf")
    _make_image_pdf(src, pages=1)
    _mock_ocr(monkeypatch)
    job = ImportJob(src_path=src, out_path=default_output_path(src), pages_to_ocr=[1])
    res = process_single_job(job, is_cancelled=lambda: True)
    assert res.status == "cancelled"
    assert not os.path.exists(job.out_path)


# 14a - rotace 0 a 180 podporovany
@pytest.mark.parametrize("rotation", [0, 180])
def test_rotation_supported(tmp_path, monkeypatch, rotation):
    src = str(tmp_path / f"rot{rotation}.pdf")
    doc = fitz.open()
    page = _add_image_page(doc, "sken")
    page.set_rotation(rotation)
    _save_doc(doc, src)
    _mock_ocr(monkeypatch, text="Rotace testovaci text")
    out = default_output_path(src)
    res = process_single_job(ImportJob(src_path=src, out_path=out, pages_to_ocr=[1]))
    assert res.status == "ok", res.message
    texts, _ = _page_texts(out)
    assert "Rotace testovaci text" in texts[0]


# 14b - rotace 90/270 bezpecne odmitnuty, zadny klamny vystup
@pytest.mark.parametrize("rotation", [90, 270])
def test_rotation_unsupported_refused(tmp_path, monkeypatch, rotation):
    src = str(tmp_path / f"rot{rotation}.pdf")
    doc = fitz.open()
    page = _add_image_page(doc, "sken")
    page.set_rotation(rotation)
    _save_doc(doc, src)
    _mock_ocr(monkeypatch)
    out = default_output_path(src)
    res = process_single_job(ImportJob(src_path=src, out_path=out, pages_to_ocr=[1]))
    assert res.status == "failed"
    assert "rotac" in res.message.lower()
    assert not os.path.exists(out)


# 15 - bez bbox: text neni tiho zahozen
def test_missing_bbox_not_lost(tmp_path, monkeypatch):
    src = str(tmp_path / "sken.pdf")
    _make_image_pdf(src, pages=1)

    def fake(images, engine, lang_raw, diacritics_enabled=False,
             progress_cb=None, is_cancelled=None):
        return [[OcrResult(text="viditelny text", bbox=_word_bbox(100, 100, 400, 60)),
                 OcrResult(text="text bez bboxu musi prezit", bbox=None)]]
    monkeypatch.setattr(pdf_import, "ocr_images_sync", fake)
    out = default_output_path(src)
    res = process_single_job(ImportJob(src_path=src, out_path=out, pages_to_ocr=[1]))
    assert res.status == "ok", res.message
    texts, _ = _page_texts(out)
    assert "text bez bboxu musi prezit" in texts[0]


# 16 - overflow na malem bboxu: text neni tiho zahozen
def test_overflow_small_bbox_not_lost(tmp_path, monkeypatch):
    src = str(tmp_path / "sken.pdf")
    _make_image_pdf(src, pages=1)
    long_text = "toto je záměrně dlouhý text který se nevejde do malinkého bboxu"

    def fake(images, engine, lang_raw, diacritics_enabled=False,
             progress_cb=None, is_cancelled=None):
        return [[OcrResult(text=long_text, bbox=_word_bbox(50, 50, 30, 12))]]
    monkeypatch.setattr(pdf_import, "ocr_images_sync", fake)
    out = default_output_path(src)
    stats_holder = {}

    orig_overlay = pdf_import.overlay_ocr_layer
    def spy(*a, **k):
        stats_holder.update(orig_overlay(*a, **k))
        return stats_holder
    monkeypatch.setattr(pdf_import, "overlay_ocr_layer", spy)
    res = process_single_job(ImportJob(src_path=src, out_path=out, pages_to_ocr=[1]))
    assert res.status == "ok", res.message
    texts, _ = _page_texts(out)
    flat = " ".join(texts[0].split())
    for word in ["záměrně", "dlouhý", "bboxu"]:
        assert word in flat
    assert stats_holder.get("failed", 0) == 0


# validace odhali poskozeny vystup (primo)
def test_validation_detects_problems(tmp_path, monkeypatch):
    src = str(tmp_path / "sken.pdf")
    _make_image_pdf(src, pages=1)
    _mock_ocr(monkeypatch, text="ocekavany text validace")
    out = default_output_path(src)
    res = process_single_job(ImportJob(src_path=src, out_path=out, pages_to_ocr=[1]))
    assert res.status == "ok"
    # spatny pocet stran
    assert validate_imported_pdf(out, 5, {}, {}) != []
    # chybejici ocekavane slovo
    bad_ocr = {1: [OcrResult(text="neco uplne jineho", bbox=_word_bbox(1, 1, 50, 20))]}
    assert validate_imported_pdf(out, 1, bad_ocr, {}) != []
    # OK pripad projde
    good_ocr = {1: [OcrResult(text="ocekavany text validace", bbox=_word_bbox(100, 100, 800, 80))]}
    assert validate_imported_pdf(out, 1, good_ocr, {}) == []


# worker: sekvencne, chyba jednoho neukonci davku (primé run bez event loopu)
def test_worker_sequential_batch(tmp_path, monkeypatch):
    try:
        from PyQt6.QtCore import QCoreApplication
        app = QCoreApplication.instance() or QCoreApplication([])
    except Exception:
        app = None
    good = str(tmp_path / "dobre.pdf")
    _make_image_pdf(good, pages=1)
    bad = str(tmp_path / "spatne.pdf")
    with open(bad, "wb") as f:
        f.write(b"garbage-not-a-pdf" * 50)
    _mock_ocr(monkeypatch, text="davkovy text")
    jobs = [
        ImportJob(src_path=bad, out_path=default_output_path(bad), pages_to_ocr=[1]),
        ImportJob(src_path=good, out_path=default_output_path(good), pages_to_ocr=[1]),
    ]
    worker = PdfImportWorker(jobs)
    collected = []
    worker.job_finished.connect(collected.append)
    finished = []
    worker.batch_finished.connect(lambda r: finished.extend(r))
    worker.run()  # synchronne ve vlakne testu (zadne GUI)
    assert len(finished) == 2
    assert finished[0].status == "failed"
    assert finished[1].status == "ok"
    assert len(collected) == 2
