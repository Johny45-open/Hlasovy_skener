"""Testy PDF exportu s potvrzeným textem včetně dělení a pokračovacích stránek.

Kontrakt (revize 3):
- Změněná stránka: extrakce obsahuje potvrzený text, starý text chybí (žádná duplicita).
- Nezměněná stránka: původní word-level vrstva beze změny.
- Dlouhý text: bezztrátové dělení na díly; první díl na zdrojové stránce
  (neviditelná vrstva), další na pokračovacích stránkách (viditelný text
  pod nadpisem "Pokračování textu stránky N") hned za zdrojovou stránkou.
- Výsledek se ukládá atomicky (tmp -> ověření -> replace); při selhání
  cílový soubor zůstane bitově nezměněný a dočasné soubory se uklidí.
- Obrazová vrstva zachována; render_mode=3 (zdroj) / 0 (pokračování), fontsize>0.
"""
import glob
import os
import tempfile

import fitz
import pytest
from PIL import Image, ImageDraw

import Skener
from ocr_engine import OcrResult
from Skener import build_searchable_pdf


DPI = 300
FONTFILE = Skener._resolve_ocr_pdf_fontfile()
A4_PX = (2480, 3508)  # A4 při 300 dpi = 595x842 pt


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


def _norm(s):
    return " ".join(s.split())


def _dense(s):
    """Text bez bílých znaků – porovnání úplnosti extrakce (MuPDF zalamuje
    i dovnitř nerozdělitelného slova, což není ztráta obsahu)."""
    return "".join(s.split())


def _group_text(path, pages, heading=None):
    """Text skupiny (zdrojová stránka + pokračování) bez bílých znaků."""
    doc = fitz.open(path)
    try:
        parts = []
        for idx in pages:
            t = doc[idx].get_text()
            if heading and idx != pages[0]:
                first, sep, rest = t.partition("\n")
                assert first.strip() == heading, (
                    f"na stránce {idx + 1} chybí nadpis {heading!r}")
                t = rest if sep else ""
            parts.append(t)
        return _dense("".join(parts))
    finally:
        doc.close()


def _tmp_leftovers(directory):
    return glob.glob(os.path.join(directory, "*.tmp"))


def _page_size_a4():
    img = _make_image(*A4_PX)
    return (img.width / DPI * 72, img.height / DPI * 72)


def _measure_newlines_capacity(page_size, fontsize=11.0):
    """Změří kapacitu prázdné oblasti zdrojové stránky v počtu zalomení."""
    w, h = page_size
    margin = Skener._confirmed_margin(w, h)
    rect = fitz.Rect(margin, margin, w - margin, h - margin)

    def fits(n):
        return Skener._probe_confirmed_fit(
            page_size, rect, "\n" * n, fontsize,
            Skener._OCR_PDF_FONTNAME, FONTFILE, 3)

    lo = 1
    while fits(lo * 2):
        lo *= 2
    hi = lo * 2
    while lo + 1 < hi:
        mid = (lo + hi) // 2
        if fits(mid):
            lo = mid
        else:
            hi = mid
    return lo


def _assert_stats_invariants(stats):
    assert stats["output_pages"] == stats["pages"] + stats["continuation"]
    assert stats["continuation"] >= len(stats["continued_pages"])


# --- Stávající kontrakt (beze změny chování) ---


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
        assert stats["continuation"] == 0
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
        _assert_stats_invariants(stats)
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


def test_whitespace_only_confirmed_text_is_empty_no_old_text():
    # Prázdný (bílý) potvrzený text nesmí být nahrazen starým OCR textem.
    img = _make_image()
    ocr = [[OcrResult(text="Starý OCR text", bbox=_word_bbox(50, 50, 400, 60))]]
    path, stats = _export([img], ocr, confirmed_override=["   \n  "])
    try:
        texts, n = _page_texts(path)
        assert n == 1
        assert "Starý OCR text" not in texts[0]
        assert stats["confirmed"] == 0
        assert stats["continuation"] == 0
    finally:
        os.remove(path)


# --- Krajní vstupy: čisté odmítnutí ---


def test_confirmed_overflow_tiny_page_rejected_cleanly():
    # Drobná stránka, na kterou se nevejde ani nadpis pokračování:
    # export se musí čistě odmítnout (RuntimeError), žádné neúplné PDF.
    img = _make_image(w=200, h=200)
    ocr = [[OcrResult(text="kratke", bbox=_word_bbox(10, 10, 100, 30))]]
    long_text = " ".join(["přeplněný"] * 2000)
    directory = tempfile.mkdtemp()
    target = os.path.join(directory, "vystup.pdf")
    try:
        with pytest.raises(RuntimeError) as excinfo:
            build_searchable_pdf([img], ocr, DPI, target,
                                 confirmed_override=[long_text])
        assert "PDF nebylo vytvořeno" in str(excinfo.value)
        assert "nevejde" in str(excinfo.value)
        # Žádné neúplné PDF: soubor nesmí vzniknout.
        assert not os.path.exists(target)
        assert not _tmp_leftovers(directory)
    finally:
        import shutil
        shutil.rmtree(directory, ignore_errors=True)


def test_progress_guard_no_infinite_loop(monkeypatch):
    # Simulace zacyklení: měření vždy tvrdí "nevejde se". Algoritmus musí
    # skončit jasnou chybou, ne smyčkou, a cíl nesmí být poškozen.
    monkeypatch.setattr(Skener, "_probe_confirmed_fit", lambda *a, **k: False)
    img = _make_image()
    ocr = [[OcrResult(text="stary", bbox=_word_bbox(50, 50, 400, 60))]]
    directory = tempfile.mkdtemp()
    target = _make_target(directory)
    try:
        with pytest.raises(RuntimeError):
            build_searchable_pdf([img], ocr, DPI, target,
                                 confirmed_override=["Ahoj světe"])
        with open(target, "rb") as f:
            assert f.read() == b"PUVODNI-EXISTUJICI-PDF"
        assert not _tmp_leftovers(directory)
    finally:
        import shutil
        shutil.rmtree(directory, ignore_errors=True)


# --- Dělení dlouhého textu: bezztrátový splitter (unit) ---


@pytest.mark.parametrize("text", [
    "Příliš žluťoučký kůň.\n\n\n\nDalší odstavec s diakritikou ěščřžýáíé.",
    ("velmi dlouhý řádek s mezerami a diakritikou řážky české. " * 300),
    ("x" * 2000),
    ("A" + " " * 10 + "B\n" * 400 + "konec"),
    "\n" * 300,
    ("slovo " * 500 + "\n\n" + "x" * 1500 + "\n\n" + "\n" * 120 + "finale"),
])
def test_splitter_roundtrip_lossless(text):
    chunks, _geom = Skener._plan_confirmed_chunks(
        _page_size_a4(), text, Skener._OCR_PDF_FONTNAME, FONTFILE, 1)
    assert chunks, "splitter musí vrátit alespoň jeden kus"
    assert "".join(c["text"] for c in chunks) == text


def test_splitter_unbreakable_word_char_fallback(monkeypatch):
    # MuPDF běžně zlomí i dlouhé slovo uprostřed, takže znakový fallback
    # vynutíme: měření povolí jen kandidáty do 100 znaků (simulace stránky,
    # kde se slovo bez dělení nevejde).
    orig = Skener._probe_confirmed_fit

    def limited(page_size, rect, text, *a, **k):
        if text is not None and len(text) > 100:
            return False
        return orig(page_size, rect, text, *a, **k)

    monkeypatch.setattr(Skener, "_probe_confirmed_fit", limited)
    word = "x" * 2000
    chunks, _ = Skener._plan_confirmed_chunks(
        _page_size_a4(), word, Skener._OCR_PDF_FONTNAME, FONTFILE, 1)
    assert "".join(c["text"] for c in chunks) == word
    assert len(chunks) >= 20, "slovo delší než kapacita musí být rozděleno"
    # Bezztrátovost roundtripu i při znakovém dělení.
    assert all(len(c["text"]) <= 100 for c in chunks)


def test_splitter_long_white_token_split_lossless():
    # Bílý blok delší než prázdná stránka i při minimálním písmu 4 pt
    # (kapacita změřena měřicím algoritmem, ne odhadem).
    page_size = _page_size_a4()
    cap = _measure_newlines_capacity(page_size, fontsize=4.0)
    assert cap > 1
    text = "\n" * (cap + 40)
    chunks, _ = Skener._plan_confirmed_chunks(
        page_size, text, Skener._OCR_PDF_FONTNAME, FONTFILE, 1)
    assert "".join(c["text"] for c in chunks) == text
    assert len(chunks) >= 2


# --- Přetečení: pokračovací stránky (end-to-end) ---


def test_short_text_single_page_no_continuation():
    img = _make_image(*A4_PX)
    ocr = [[OcrResult(text="stary", bbox=_word_bbox(50, 50, 400, 60))]]
    text = "Krátký potvrzený text, který se vejde na jednu stránku."
    path, stats = _export([img], ocr, confirmed_override=[text])
    try:
        texts, n = _page_texts(path)
        assert n == 1
        assert stats["continuation"] == 0
        assert stats["output_pages"] == 1
        _assert_stats_invariants(stats)
        assert _norm(texts[0]) == _norm(text)
    finally:
        os.remove(path)


def test_long_text_overflows_to_continuation():
    img = _make_image(*A4_PX)
    ocr = [[OcrResult(text="stary OCR", bbox=_word_bbox(50, 50, 400, 60))]]
    # ~69 tis. znaků – více než kapacita A4 i při minimálním písmu 4 pt.
    text = "Řádek s českou diakritikou, přeplněná stránka. " * 1500
    path, stats = _export([img], ocr, confirmed_override=[text])
    try:
        _assert_stats_invariants(stats)
        assert stats["confirmed"] == 1
        assert stats["continuation"] >= 1, "dlouhý text musí přetéct"
        texts, n = _page_texts(path)
        assert n == stats["output_pages"] == 1 + stats["continuation"]
        got = _group_text(path, list(range(n)),
                          heading="Pokračování textu stránky 1")
        assert got == _dense(text)
        assert "staryOCR" not in got
    finally:
        os.remove(path)


def test_very_long_text_thousands_words_completely_saved():
    img = _make_image(*A4_PX)
    ocr = [[OcrResult(text="stary", bbox=_word_bbox(50, 50, 400, 60))]]
    text = ("Opravený řádek s diakritikou ěščřžýáíé úůďťň. " * 2500)
    path, stats = _export([img], ocr, confirmed_override=[text])
    try:
        _assert_stats_invariants(stats)
        assert stats["continuation"] >= 1
        got = _group_text(path, list(range(stats["output_pages"])),
                          heading="Pokračování textu stránky 1")
        exp = _dense(text)
        assert got == exp
        # Pořadí: začátek i konec textu na správných místech.
        assert got[:60] == exp[:60]
        assert got[-60:] == exp[-60:]
    finally:
        os.remove(path)


def test_continuation_heading_present_and_excluded():
    img = _make_image(*A4_PX)
    ocr = [[OcrResult(text="stary", bbox=_word_bbox(50, 50, 400, 60))]]
    text = "Začátek potvrzeného textu. " * 2000
    path, stats = _export([img], ocr, confirmed_override=[text])
    try:
        assert stats["continuation"] >= 1
        heading = "Pokračování textu stránky 1"
        doc = fitz.open(path)
        try:
            for idx in range(1, stats["output_pages"]):
                first, sep, rest = doc[idx].get_text().partition("\n")
                assert first.strip() == heading
                assert heading not in rest  # nadpis není součástí obsahu
        finally:
            doc.close()
        # Nadpis není součástí potvrzeného textu ani po odstranění mezer.
        got = _group_text(path, list(range(stats["output_pages"])),
                          heading=heading)
        assert _dense(heading) not in got
        assert got == _dense(text)
    finally:
        os.remove(path)


def test_multi_page_with_continuation_assignment():
    imgs = [_make_image(*A4_PX) for _ in range(3)]
    ocr = [
        [OcrResult(text="alfa obsah", bbox=_word_bbox(50, 50, 400, 60))],
        [OcrResult(text="beta obsah", bbox=_word_bbox(50, 50, 400, 60))],
        [OcrResult(text="gama obsah", bbox=_word_bbox(50, 50, 400, 60))],
    ]
    long_beta = "Opravená beta stránka s hodně textem a diakritikou řščřž. " * 1500
    path, stats = _export(imgs, ocr,
                          confirmed_override=[None, long_beta, None])
    try:
        _assert_stats_invariants(stats)
        assert stats["continuation"] >= 1
        assert stats["continued_pages"] == [1]
        texts, n = _page_texts(path)
        assert n == 3 + stats["continuation"]
        assert "alfa obsah" in texts[0]
        assert "gama obsah" in texts[-1]
        group_pages = list(range(1, 1 + 1 + stats["continuation"]))
        got = _group_text(
            path, group_pages, heading="Pokračování textu stránky 2")
        assert got == _dense(long_beta)
        # Starý OCR text nesmí být nikde ve skupině změněné stránky.
        for idx in group_pages:
            assert "beta obsah" not in texts[idx]
    finally:
        os.remove(path)


def test_stats_invariants_mixed_document():
    imgs = [_make_image(*A4_PX) for _ in range(2)]
    ocr = [
        [OcrResult(text="jedna", bbox=_word_bbox(50, 50, 400, 60))],
        [OcrResult(text="dva", bbox=_word_bbox(50, 50, 400, 60))],
    ]
    long_text = "Velmi dlouhý text druhé stránky. " * 2000
    path, stats = _export(imgs, ocr,
                          confirmed_override=["Krátká změna", long_text])
    try:
        _assert_stats_invariants(stats)
        assert stats["pages"] == 2
        assert stats["confirmed"] == 2
        assert stats["output_pages"] == 2 + stats["continuation"]
        assert stats["continued_pages"] == [1]
    finally:
        os.remove(path)


# --- Integrační regresní test (kombinovaný vstup) ---


def test_integration_complex_combination():
    page_size = _page_size_a4()
    cap = _measure_newlines_capacity(page_size)
    white_block = "\n" * max(2, cap // 2)
    assert 0 < len(white_block) < cap, "bílý blok musí být menší než prázdná stránka"

    cz = "Úvodní odstavec s českou diakritikou: ěščřžýáíéúůďťň. "
    long_line = ("velmi dlouhý řádek s mezerami a diakritikou řážky. " * 1500)
    unbreakable = "x" * 2000
    tail = "Závěrečná věta po bílém bloku."
    text = (cz + "\n" + long_line + "\n\n" + unbreakable
            + "\n\n\n" + white_block + tail)

    # 1) Přesná rekonstrukce ze splitteru.
    chunks, _ = Skener._plan_confirmed_chunks(
        page_size, text, Skener._OCR_PDF_FONTNAME, FONTFILE, 1)
    assert "".join(c["text"] for c in chunks) == text

    # 2) Export a kontrola každé skupiny v znovu otevřeném PDF.
    img = _make_image(*A4_PX)
    ocr = [[OcrResult(text="stary OCR text", bbox=_word_bbox(50, 50, 400, 60))]]
    path, stats = _export([img], ocr, confirmed_override=[text])
    try:
        _assert_stats_invariants(stats)
        assert stats["continuation"] >= 2, "kombinovaný vstup musí přetéct"
        heading = "Pokračování textu stránky 1"
        got = _group_text(path, list(range(stats["output_pages"])),
                          heading=heading)
        exp = _dense(text)
        # Úplnost, pořadí (začátek/konec) a oddělení nadpisů.
        assert got == exp
        assert got[:60] == exp[:60]
        assert got[-40:] == exp[-40:]
        assert _dense(heading) not in got
        assert "staryOCRtext" not in got
    finally:
        os.remove(path)


# --- Atomické ukládání a ochrana cílového souboru ---


def _make_target(directory):
    target = tempfile.NamedTemporaryFile(
        dir=directory, delete=False, suffix=".pdf")
    target.write(b"PUVODNI-EXISTUJICI-PDF")
    target.close()
    return target.name


def test_target_preserved_on_save_failure(monkeypatch):
    def boom(self, *a, **k):
        raise RuntimeError("simulace selhání zápisu")

    monkeypatch.setattr(fitz.Document, "save", boom)
    img = _make_image()
    ocr = [[OcrResult(text="stary", bbox=_word_bbox(50, 50, 400, 60))]]
    directory = tempfile.mkdtemp()
    target = _make_target(directory)
    try:
        with pytest.raises(RuntimeError):
            build_searchable_pdf([img], ocr, DPI, target,
                                 confirmed_override=["nový text"])
        with open(target, "rb") as f:
            assert f.read() == b"PUVODNI-EXISTUJICI-PDF"
        assert not _tmp_leftovers(directory)
    finally:
        os.remove(target)
        os.rmdir(directory)


def test_target_preserved_on_verification_failure(monkeypatch):
    def boom(vdoc, groups):
        raise RuntimeError("simulace selhání ověření obsahu")

    monkeypatch.setattr(Skener, "_verify_confirmed_groups", boom)
    img = _make_image()
    ocr = [[OcrResult(text="stary", bbox=_word_bbox(50, 50, 400, 60))]]
    directory = tempfile.mkdtemp()
    target = _make_target(directory)
    try:
        with pytest.raises(RuntimeError):
            build_searchable_pdf([img], ocr, DPI, target,
                                 confirmed_override=["nový text"])
        with open(target, "rb") as f:
            assert f.read() == b"PUVODNI-EXISTUJICI-PDF"
        assert not _tmp_leftovers(directory)
    finally:
        os.remove(target)
        os.rmdir(directory)


def test_target_preserved_on_replace_failure(monkeypatch):
    def boom(src, dst):
        raise OSError("simulace selhání nahrazování cíle")

    monkeypatch.setattr(os, "replace", boom)
    img = _make_image()
    ocr = [[OcrResult(text="stary", bbox=_word_bbox(50, 50, 400, 60))]]
    directory = tempfile.mkdtemp()
    target = _make_target(directory)
    try:
        with pytest.raises(OSError):
            build_searchable_pdf([img], ocr, DPI, target,
                                 confirmed_override=["nový text"])
        with open(target, "rb") as f:
            assert f.read() == b"PUVODNI-EXISTUJICI-PDF"
        assert not _tmp_leftovers(directory)
    finally:
        os.remove(target)
        os.rmdir(directory)
