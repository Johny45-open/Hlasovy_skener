"""Regresní testy bezpečné diakritiky (krok 3).

- Žádná automatická významová změna (byt → být a audit celého slovníku).
- Běžné doplnění diakritiky stále funguje.
- Case, interpunkce, čísla, mezery a odřádkování zachovány.
- Transformers backend není implementován (vždy slovník) – dokumentováno testem.
"""
import diacritics as dia
from diacritics import (
    _AMBIGUOUS_STRIPPED,
    _CORRECT_SET,
    _DIACRITICS_DICT,
    _transformers_diacritize,
    diacritize_text,
    get_diacritizer,
    should_diacritize,
)


def test_byt_never_changed_to_byt():
    assert diacritize_text("byt") == "byt"
    assert diacritize_text("Byt") == "Byt"
    assert diacritize_text("BYT") == "BYT"
    assert diacritize_text("Mám byt.") == "Mám byt."
    assert diacritize_text("Bydlím v bytě. Mám byt. Chci být doma.") == \
        "Bydlím v bytě. Mám byt. Chci být doma."


def test_already_correct_byt_stays():
    assert diacritize_text("být") == "být"
    assert diacritize_text("Chci být doma.") == "Chci být doma."


def test_byt_is_ambiguous():
    assert "byt" in _AMBIGUOUS_STRIPPED
    assert _DIACRITICS_DICT.get("byt") is None


def test_no_dict_key_maps_to_another_valid_word():
    """Audit celého slovníku: stripped klíč nesmí přepisovat jiné platné slovo."""
    offenders = {k: v for k, v in _DIACRITICS_DICT.items()
                 if k in _CORRECT_SET and k != v}
    assert offenders == {}, f"významové kolize: {offenders}"


def test_no_ambiguous_key_in_dict():
    assert not (set(_DIACRITICS_DICT) & _AMBIGUOUS_STRIPPED)


def test_ordinary_diacritics_still_work():
    assert diacritize_text("krasne pocasi") == "krásné počasí"
    assert diacritize_text("prilis zlutoucky kun upel dabelske ody") == \
        "příliš žluťoučký kůň úpěl ďábelské ódy"


def test_case_punctuation_numbers_whitespace_preserved():
    assert diacritize_text("KRASNE POCASI!") == "KRÁSNÉ POČASÍ!"
    assert diacritize_text("Krasne, pocasi...") == "Krásné, počasí..."
    assert diacritize_text("cislo 123 a 45Test") == "číslo 123 a 45Test"
    assert diacritize_text("prvni\n  druhy\tstranka") == "první\n  druhý\tstránka"


def test_unknown_words_unchanged():
    assert diacritize_text("xyzqwx blaf") == "xyzqwx blaf"


def test_page_header_survives():
    text = "--- Stránka 1 ---\nkrasne pocasi\n"
    assert diacritize_text(text) == "--- Stránka 1 ---\nkrásné počasí\n"


def test_should_diacritize_only_ces_and_enabled():
    assert should_diacritize("ces (Čeština)", True) is True
    assert should_diacritize("ces", True) is True
    assert should_diacritize("ces (Čeština)", False) is False
    assert should_diacritize("eng (Angličtina)", True) is False
    assert should_diacritize("", True) is False


def test_diacritizer_never_raises_and_keeps_original_on_empty():
    d = get_diacritizer()
    assert d.is_available()
    assert d.diacritize("") == ""
    assert d.diacritize("   ") == "   "


def test_no_transformers_backend_always_dictionary():
    """Dokumentuje, že BERT backend není implementován – vždy slovník, offline."""
    assert _transformers_diacritize("krasne pocasi") is None
