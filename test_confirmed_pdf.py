"""Testy opraveného PDF exportu s potvrzeným textem (krok 2).

- Změněná stránka: extrakce obsahuje potvrzený text, starý text chybí (žádná duplicita).
- Nezměněná stránka: původní word-level vrstva beze změny.
- Přiřazení ke stránce u více stránek.
- Overflow potvrzené vrstvy: RuntimeError, žádné částečné PDF.
- Obrazová vrstva zachována; render_mode=3, fontsize>0.
"""
import os
import tempfile

import fitz
import pytest
from PIL import Image, ImageDraw

from ocr_engine import OcrResult
from Skener import build_searchable_pdf


DPI = 300


def _make_image(w=1200, h=1600, text="test"):
    img = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(img)
    d.text((50, 50), text, fill="black")
    return img


def _word_bbox(x, y, w, h):
    return [[float(x), float(y)], [float(x + w), float(y)],
            [float(x + w), float(y + h)], [float(x), float(y + h)]]


def _export(images, ocr_pages, dpi=DPI, confirmed_override=None):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    tmp.close()
    stats = build_searchable_pdf(images, ocr_pages, dpi, tmp.name,
                                 confirmed_override=confirmed_override)
    return tmp.name, stats


def _page_texts(path):
    doc = fitz.open(path)
    try:
        return [p.get_text().strip() for p in doc], len(doc)
    finally:
        doc.close()


def test_changed_page_contains_confirmed_not_old():
    img = _make_image()
    ocr = [[OcrResult(text="Hello", bbox=_word_bbox(50, 50, 400, 60))]]
    path, stats = _export([img], ocr, confirmed_override=["Opraveno"])
    try:
        texts, n = _page_texts(path)
        assert n == 1
        assert "Opraveno" in texts[0]
        assert "Hello" not in texts[0]
        assert stats["confirmed"] == 1
        assert stats["failed"] == 0
    finally:
        os.remove(path)


def test_changed_page_no_duplication_single_occurrence():
    img = _make_image()
    ocr = [[OcrResult(text="stary text", bbox=_word_bbox(50, 50, 400, 60))]]
    path, _ = _export([img], ocr, confirmed_override=["nový text"])
    try:
        texts, _ = _page_texts(path)
        assert texts[0].count("nový text") == 1
        assert "stary text" not in texts[0]
    finally:
        os.remove(path)


def test_unchanged_pages_keep_wordlevel():
    img = _make_image()
    ocr = [[OcrResult(text="Hello world", bbox=_word_bbox(50, 50, 400, 60))]]
    path, stats = _export([img], ocr, confirmed_override=[None])
    try:
        texts, _ = _page_texts(path)
        assert "Hello world" in texts[0]
        assert stats["confirmed"] == 0
        assert stats["inserted"] >= 1
    finally:
        os.remove(path)


def test_legacy_call_without_override_unchanged():
    img = _make_image()
    ocr = [[OcrResult(text="legacy text", bbox=_word_bbox(50, 50, 400, 60))]]
    path, stats = _export([img], ocr)
    try:
        texts, _ = _page_texts(path)
        assert "legacy text" in texts[0]
        assert stats["confirmed"] == 0
    finally:
        os.remove(path)


def test_multi_page_mixed_assignment():
    imgs = [_make_image() for _ in range(3)]
    ocr = [
        [OcrResult(text="alfa obsah", bbox=_word_bbox(50, 50, 400, 60))],
        [OcrResult(text="beta obsah", bbox=_word_bbox(50, 50, 400, 60))],
        [OcrResult(text="gama obsah", bbox=_word_bbox(50, 50, 400, 60))],
    ]
    path, stats = _export(imgs, ocr,
                          confirmed_override=[None, "Beta opravená", None])
    try:
        texts, n = _page_texts(path)
        assert n == 3
        assert "alfa obsah" in texts[0]
        assert "Beta opravená" in texts[1]
        assert "beta obsah" not in texts[1]
        assert "gama obsah" in texts[2]
        assert stats["confirmed"] == 1
    finally:
        os.remove(path)


def test_confirmed_diacritics_preserved():
    img = _make_image()
    cz = "Příliš žluťoučký kůň úpěl ďábelské ódy"
    ocr = [[OcrResult(text="Prilis zlutoucky kun", bbox=_word_bbox(50, 50, 400, 60))]]
    path, _ = _export([img], ocr, confirmed_override=[cz])
    try:
        texts, _ = _page_texts(path)
        assert cz in texts[0]
        for ch in ["ř", "š", "ů", "ň", "ď", "ť", "ž", "ý", "á", "í", "é", "ú", "ó"]:
            if ch in cz:
                assert ch in texts[0]
    finally:
        os.remove(path)


def test_confirmed_layer_uses_full_page_not_bottom_strip(monkeypatch):
    calls = []
    orig = fitz.Page.insert_textbox

    def spy(self, rect, buffer, **kwargs):
        calls.append((rect, buffer, dict(kwargs)))
        return orig(self, rect, buffer, **kwargs)

    monkeypatch.setattr(fitz.Page, "insert_textbox", spy)
    img = _make_image()
    ocr = [[OcrResult(text="stary", bbox=_word_bbox(50, 50, 400, 60))]]
    path, _ = _export([img], ocr, confirmed_override=["nový potvrzený text"])
    try:
        assert calls, "očekáváno vložení potvrzené vrstvy"
        for rect, _buf, kw in calls:
            assert kw.get("render_mode") == 3
            assert float(kw["fontsize"]) > 0
        # Vrstva musí pokrývat podstatnou část stránky, ne úzký spodní pruh.
        rects = [c[0] for c in calls]
        tallest = max(r.y1 - r.y0 for r in rects)
        assert tallest > 200, f"vrstva je omezená na pruh ({tallest}pt)"
    finally:
        os.remove(path)


def test_confirmed_overflow_aborts_no_partial_pdf():
    # Drobná stránka + velmi dlouhý text se nesmí potichu ořezat.
    img = _make_image(w=200, h=200)
    ocr = [[OcrResult(text="kratke", bbox=_word_bbox(10, 10, 100, 30))]]
    long_text = " ".join(["přeplněný"] * 2000)
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    tmp.close()
    try:
        with pytest.raises(RuntimeError):
            build_searchable_pdf([img], ocr, DPI, tmp.name,
                                 confirmed_override=[long_text])
        # Žádné neúplné PDF nesmí zůstat (soubor buď chybí, nebo je prázdný/bez save).
        if os.path.exists(tmp.name):
            assert os.path.getsize(tmp.name) == 0
    finally:
        if os.path.exists(tmp.name):
            os.remove(tmp.name)


def test_image_layer_preserved_for_changed_page():
    img = _make_image()
    ocr = [[OcrResult(text="stary", bbox=_word_bbox(50, 50, 400, 60))]]
    path, _ = _export([img], ocr, confirmed_override=["nový"])
    try:
        doc = fitz.open(path)
        try:
            assert len(doc[0].get_images(full=True)) >= 1
        finally:
            doc.close()
    finally:
        os.remove(path)


def test_override_length_mismatch_raises():
    img = _make_image()
    ocr = [[OcrResult(text="a", bbox=_word_bbox(50, 50, 400, 60))]]
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    tmp.close()
    try:
        with pytest.raises(ValueError):
            build_searchable_pdf([img], ocr, DPI, tmp.name,
                                 confirmed_override=[None, None])
    finally:
        if os.path.exists(tmp.name):
            os.remove(tmp.name)


def test_emptied_page_has_no_old_text():
    img = _make_image()
    ocr = [[OcrResult(text="Starý text", bbox=_word_bbox(50, 50, 400, 60))]]
    path, _ = _export([img], ocr, confirmed_override=[""])
    try:
        texts, _ = _page_texts(path)
        assert "Starý text" not in texts[0]
    finally:
        os.remove(path)
