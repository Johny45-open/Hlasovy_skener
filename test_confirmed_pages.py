"""Testy potvrzeného textu po stránkách (krok 1).

parse_confirmed_pages_strict: striktní validace struktury před potvrzením.
build_confirmed_pdf_plan: rozhodnutí word-level vs. explicitní vrstva.
"""
import pytest

from ocr_engine import OcrResult
from Skener import (
    parse_confirmed_pages_strict,
    build_confirmed_pdf_plan,
)


def _bbox(x=50, y=50, w=400, h=60):
    return [[float(x), float(y)], [float(x + w), float(y)],
            [float(x + w), float(y + h)], [float(x), float(y + h)]]


# --- strict parser: jedna stránka ---

def test_single_page_no_headers_is_page_one():
    pages = parse_confirmed_pages_strict("Opravený text bez headeru", 1)
    assert pages == ["Opravený text bez headeru"]


def test_single_page_with_header_one():
    pages = parse_confirmed_pages_strict("--- Stránka 1 ---\nAhoj světe\n", 1)
    assert pages == ["Ahoj světe"]


def test_single_page_wrong_header_number_rejected():
    assert parse_confirmed_pages_strict("--- Stránka 2 ---\ntext", 1) is None


def test_single_page_text_before_header_rejected():
    assert parse_confirmed_pages_strict("poznámka\n--- Stránka 1 ---\ntext", 1) is None


def test_single_page_two_headers_rejected():
    assert parse_confirmed_pages_strict(
        "--- Stránka 1 ---\na\n--- Stránka 1 ---\nb", 1) is None


# --- strict parser: více stránek ---

def test_multi_page_ok():
    text = "--- Stránka 1 ---\nAlfa\n\n--- Stránka 2 ---\nBeta\n"
    assert parse_confirmed_pages_strict(text, 2) == ["Alfa", "Beta"]


def test_multi_page_missing_header_rejected():
    text = "--- Stránka 1 ---\nAlfa\n\n--- Stránka 3 ---\nGama\n"
    assert parse_confirmed_pages_strict(text, 2) is None


def test_multi_page_wrong_order_rejected():
    text = "--- Stránka 2 ---\nBeta\n\n--- Stránka 1 ---\nAlfa\n"
    assert parse_confirmed_pages_strict(text, 2) is None


def test_multi_page_no_headers_rejected_no_silent_merge():
    # Text bez headerů pro 2 stránky se NESMÍ potichu přesunout na stranu 1.
    assert parse_confirmed_pages_strict("souvislý text bez headerů", 2) is None


def test_multi_page_text_before_first_header_rejected():
    text = "úvod\n--- Stránka 1 ---\nAlfa\n\n--- Stránka 2 ---\nBeta\n"
    assert parse_confirmed_pages_strict(text, 2) is None


def test_multi_page_extra_header_rejected():
    text = ("--- Stránka 1 ---\nA\n--- Stránka 2 ---\nB\n"
            "--- Stránka 3 ---\nC\n")
    assert parse_confirmed_pages_strict(text, 2) is None


def test_zero_total_rejected():
    assert parse_confirmed_pages_strict("cokoliv", 0) is None
    assert parse_confirmed_pages_strict("", 0) is None


def test_inner_linebreaks_preserved():
    text = "--- Stránka 1 ---\nřádek jedna\nřádek dva\n"
    pages = parse_confirmed_pages_strict(text, 1)
    assert pages == ["řádek jedna\nřádek dva"]


# --- plan: word-level vs. potvrzená vrstva ---

def test_plan_unchanged_keeps_wordlevel():
    ocr = [[OcrResult(text="Hello world", bbox=_bbox())]]
    override, changed = build_confirmed_pdf_plan(ocr, ["Hello world"], 1)
    assert changed == [False]
    assert override == [None]


def test_plan_whitespace_only_diff_is_unchanged():
    ocr = [[OcrResult(text="Hello world", bbox=_bbox())]]
    override, changed = build_confirmed_pdf_plan(ocr, ["  Hello   world\n"], 1)
    assert changed == [False]
    assert override == [None]


def test_plan_changed_page_uses_confirmed():
    ocr = [[OcrResult(text="Hello", bbox=_bbox())]]
    override, changed = build_confirmed_pdf_plan(ocr, ["Opraveno"], 1)
    assert changed == [True]
    assert override == ["Opraveno"]


def test_plan_mixed_pages():
    ocr = [
        [OcrResult(text="Alfa", bbox=_bbox())],
        [OcrResult(text="Beta", bbox=_bbox())],
        [OcrResult(text="Gama", bbox=_bbox())],
    ]
    override, changed = build_confirmed_pdf_plan(
        ocr, ["Alfa", "Beta opravená", "Gama"], 3)
    assert changed == [False, True, False]
    assert override == [None, "Beta opravená", None]


def test_plan_no_confirmation_is_legacy():
    ocr = [[OcrResult(text="Hello", bbox=_bbox())]]
    override, changed = build_confirmed_pdf_plan(ocr, None, 1)
    assert changed == [False]
    assert override == [None]


def test_plan_length_mismatch_is_legacy():
    ocr = [[OcrResult(text="A", bbox=_bbox())]]
    override, changed = build_confirmed_pdf_plan(ocr, ["A", "B"], 1)
    assert changed == [False]
    assert override == [None]


def test_plan_empty_both_is_unchanged():
    override, changed = build_confirmed_pdf_plan([[]], [""], 1)
    assert changed == [False]
    assert override == [None]


def test_plan_user_typed_into_empty_page_is_changed():
    override, changed = build_confirmed_pdf_plan([[]], ["Dopsaný text"], 1)
    assert changed == [True]
    assert override == ["Dopsaný text"]


def test_plan_user_emptied_page_is_changed_no_old_text_leak():
    ocr = [[OcrResult(text="Starý text", bbox=_bbox())]]
    override, changed = build_confirmed_pdf_plan(ocr, [""], 1)
    assert changed == [True]
    assert override == [""]
