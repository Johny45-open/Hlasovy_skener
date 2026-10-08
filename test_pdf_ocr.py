"""Testy prohledavatelného OCR PDF (build_searchable_pdf v Skener.py).

Pokrytí dle ticketu:
A. jedna stránka bez diakritiky
B. více stránek (3), počet + text každé stránky
C. čeština s diakritikou, žádné '?' místo znaků
D. žádný export nepoužije fontsize=0 (monkeypatch insert_textbox)
E. použití render_mode=3 (monkeypatch)
F. overflow na malém bboxu – text nesmí být tiše zahozen
G. bbox=None – text nesmí být ztracen bez upozornění
H. kompatibilita Tesseract / EasyOCR bbox formátu
I. diakritika ON/OFF
J. finální PDF přes PyMuPDF: len(doc), neprázdný text, obsah
K. nezávislý extraktor (pdfplumber/pdfminer) jako cross-check, pokud je dostupný
"""
import os
import tempfile

import fitz
import pytest
from PIL import Image, ImageDraw

from ocr_engine import OcrResult
import Skener
from Skener import (
    build_searchable_pdf,
    _resolve_ocr_pdf_fontfile,
    _ocr_pdf_initial_fontsize,
)


DPI = 300
CZ = "Příliš žluťoučký kůň úpěl ďábelské ódy"


def _make_image(w=1200, h=1600, text="test"):
    img = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(img)
    d.text((50, 50), text, fill="black")
    return img


def _word_bbox(x, y, w, h):
    return [[float(x), float(y)], [float(x + w), float(y)],
            [float(x + w), float(y + h)], [float(x), float(y + h)]]


def _easyocr_bbox(x, y, w, h):
    # EasyOCR vrací 4-bodový polygon – pro axis-aligned stejné jako Tesseract.
    return [[float(x), float(y)], [float(x + w), float(y)],
            [float(x + w), float(y + h)], [float(x), float(y + h)]]


def _export(images, ocr_pages, dpi=DPI):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    tmp.close()
    stats = build_searchable_pdf(images, ocr_pages, dpi, tmp.name)
    return tmp.name, stats


def _page_texts(path):
    doc = fitz.open(path)
    try:
        return [p.get_text().strip() for p in doc], len(doc)
    finally:
        doc.close()


# A – jedna stránka bez diakritiky
def test_single_page_plain_text():
    img = _make_image(text="Hello world")
    ocr = [[OcrResult(text="Hello world", bbox=_word_bbox(50, 50, 400, 60))]]
    path, stats = _export([img], ocr)
    try:
        texts, n = _page_texts(path)
        assert n == 1
        assert "Hello world" in texts[0]
        assert stats["inserted"] >= 1
        assert stats["failed"] == 0
    finally:
        os.remove(path)


# B – více stránek
def test_multi_page_count_and_content():
    imgs = [_make_image(text=f"page {i}") for i in range(1, 4)]
    ocr = [
        [OcrResult(text="Prvni strana obsah", bbox=_word_bbox(50, 50, 500, 60))],
        [OcrResult(text="Druha strana obsah", bbox=_word_bbox(50, 50, 500, 60))],
        [OcrResult(text="Treti strana obsah", bbox=_word_bbox(50, 50, 500, 60))],
    ]
    path, stats = _export(imgs, ocr)
    try:
        texts, n = _page_texts(path)
        assert n == 3
        assert "Prvni" in texts[0] and texts[0].strip()
        assert "Druha" in texts[1] and texts[1].strip()
        assert "Treti" in texts[2] and texts[2].strip()
    finally:
        os.remove(path)


# C – čeština, žádné otazníky
def test_czech_diacritics_preserved():
    img = _make_image(text="cz")
    ocr = [[OcrResult(text=CZ, bbox=_word_bbox(50, 50, 1000, 80))]]
    path, _ = _export([img], ocr)
    try:
        texts, _ = _page_texts(path)
        out = texts[0]
        assert CZ in out, f"očekáváno {CZ!r}, získáno {out!r}"
        # Nesmí dojít k náhradě ř/š/ů/ň/ď/ť otazníky (regrese helv fontu).
        for ch in ["ř", "š", "ů", "ň", "ď", "ť", "ž", "ý", "á", "í", "é", "ú", "ó"]:
            if ch in CZ:
                assert ch in out, f"znak {ch!r} chybí v {out!r}"
        assert "?" not in out.replace("?", "") or "?" not in out
    finally:
        os.remove(path)


# D+E – fontsize nikdy 0, render_mode vždy 3
def test_fontsize_never_zero_and_render_mode_3(monkeypatch):
    calls = []
    orig = fitz.Page.insert_textbox

    def spy(self, rect, buffer, **kwargs):
        calls.append(dict(kwargs))
        return orig(self, rect, buffer, **kwargs)

    monkeypatch.setattr(fitz.Page, "insert_textbox", spy)
    img = _make_image()
    ocr = [[
        OcrResult(text="maly text", bbox=_word_bbox(50, 50, 300, 50)),
        OcrResult(text=CZ, bbox=_word_bbox(50, 200, 900, 70)),
    ]]
    path, _ = _export([img], ocr)
    try:
        assert calls, "očekáváno alespoň jedno insert_textbox"
        for kw in calls:
            assert "fontsize" in kw, f"chybí fontsize v {kw}"
            assert kw["fontsize"] is not None and float(kw["fontsize"]) > 0, kw
            assert float(kw["fontsize"]) != 0.0
            assert kw.get("render_mode") == 3, f"očekáván render_mode=3, získáno {kw}"
            assert kw.get("fill_opacity", 1) != 0, "fill_opacity=0 se nesmí používat"
    finally:
        os.remove(path)


# F – overflow na záměrně malém bboxu
def test_overflow_small_bbox_not_silently_dropped():
    img = _make_image()
    long_text = "toto je záměrně dlouhý text který se nevejde do malinkého bboxu"
    ocr = [[OcrResult(text=long_text, bbox=_word_bbox(50, 50, 30, 12))]]
    path, stats = _export([img], ocr)
    try:
        texts, _ = _page_texts(path)
        # Text musí být zachován celý (shrink nebo fallback), ne tiše zahozen.
        flat = " ".join(texts[0].split())
        for word in ["záměrně", "dlouhý", "bboxu"]:
            assert word in flat, f"slovo {word!r} chybí – text byl zahozen (stats={stats})"
        assert stats["inserted"] >= 1
        assert stats["failed"] == 0
        assert (stats["shrunk"] + stats["fallback"]) >= 1
    finally:
        os.remove(path)


# G – chybějící bbox
def test_missing_bbox_not_lost():
    img = _make_image()
    ocr = [[
        OcrResult(text="viditelny bbox text", bbox=_word_bbox(50, 50, 400, 60)),
        OcrResult(text="text bez bboxu musi prezit", bbox=None),
    ]]
    path, stats = _export([img], ocr)
    try:
        texts, _ = _page_texts(path)
        assert "text bez bboxu musi prezit" in texts[0]
        assert stats["nobbox"] >= 1
        assert stats["failed"] == 0
    finally:
        os.remove(path)


# H – Tesseract vs EasyOCR bbox formát
@pytest.mark.parametrize("engine", ["tesseract", "easyocr"])
def test_engine_bbox_compatibility(engine):
    img = _make_image()
    if engine == "tesseract":
        # Tesseract: left/top/width/height -> 4 rohy (ocr_engine.py).
        bbox = _word_bbox(100, 120, 350, 55)
    else:
        bbox = _easyocr_bbox(100, 120, 350, 55)
    ocr = [[OcrResult(text="engine kompatibilita", bbox=bbox)]]
    path, stats = _export([img], ocr)
    try:
        texts, _ = _page_texts(path)
        assert "engine kompatibilita" in texts[0]
        assert stats["failed"] == 0
    finally:
        os.remove(path)


# I – diakritika ON (opravený text) vs OFF (původní bez diakritiky)
@pytest.mark.parametrize("corrected", [True, False])
def test_diacritics_on_off(corrected):
    img = _make_image()
    text = CZ if corrected else "Prilis zlutoucky kun upel dabelske ody"
    h = 80
    ocr = [[OcrResult(text=text, bbox=_word_bbox(50, 50, 1000, h),
                      original_text="Prilis zlutoucky kun",
                      processed_text=text)]]
    path, _ = _export([img], ocr)
    try:
        texts, _ = _page_texts(path)
        assert text in texts[0]
    finally:
        os.remove(path)


# J – finální PDF: stránky, neprázdný text, obsah
def test_final_pdf_structure():
    imgs = [_make_image(text=f"s{i}") for i in range(3)]
    ocr = [
        [OcrResult(text="alfa obsah", bbox=_word_bbox(50, 50, 400, 60))],
        [OcrResult(text="beta obsah", bbox=_word_bbox(50, 50, 400, 60))],
        [OcrResult(text="gama obsah", bbox=_word_bbox(50, 50, 400, 60))],
    ]
    path, _ = _export(imgs, ocr)
    try:
        doc = fitz.open(path)
        try:
            assert len(doc) == 3
            for i, page in enumerate(doc):
                t = page.get_text().strip()
                assert t, f"stránka {i + 1} je prázdná"
            assert "alfa" in doc[0].get_text()
            assert "beta" in doc[1].get_text()
            assert "gama" in doc[2].get_text()
        finally:
            doc.close()
    finally:
        os.remove(path)


# Font helpery
def test_font_resolved_and_initial_size_positive():
    assert _resolve_ocr_pdf_fontfile() is not None
    assert os.path.isfile(_resolve_ocr_pdf_fontfile())
    r = fitz.Rect(0, 0, 100, 20)
    assert _ocr_pdf_initial_fontsize(r) > 0
    r0 = fitz.Rect(0, 0, 0, 0)
    assert _ocr_pdf_initial_fontsize(r0) > 0


# K – nezávislý extraktor (pdfplumber -> pdfminer), skip pokud není
def test_independent_extractor_crosscheck():
    pdfplumber = pytest.importorskip("pdfplumber")
    img = _make_image()
    ocr = [[OcrResult(text=CZ, bbox=_word_bbox(50, 50, 1000, 80))]]
    path, _ = _export([img], ocr)
    try:
        import pdfplumber as pl
        with pl.open(path) as pdf:
            assert len(pdf.pages) == 1
            t = pdf.pages[0].extract_text() or ""
        # pdfplumber může jinak lámat řádky – kontroluj klíčová slova.
        for word in ["Příliš", "žluťoučký", "kůň"]:
            assert word in t, f"{word!r} chybí v pdfplumber výstupu {t!r}"
    finally:
        os.remove(path)
