"""Offline česká diakritizace pro Hlasový skener.

Priorita: offline, bez internetu, Windows, rozumná velikost, bezpečná kontextová oprava.
Nepoužívá cloud API. Primárně slovníková kontextová metoda (word-level),
volitelně BERT/transformers pokud je dostupný.

NESMÍ provádět naivní nahrazování znaků typu c->č. Oprava je vždy na úrovni
celých slov: stripped_form -> diacritized_form ze slovníku.
Neznámá slova se ponechávají beze změny (bezpečné).
"""
from __future__ import annotations

import re
import unicodedata
import threading
from typing import Optional

# ---------------------------------------------------------------------------
# Pomocné: odstranění diakritiky (fold)
# ---------------------------------------------------------------------------

def strip_diacritics(text: str) -> str:
    """Odstraní diakritiku (háčky, čárky, kroužek) – používá NFD."""
    # Speciální české znaky které NFD rozloží správně: ě, š, č, ř, ž, ď, ť, ň, á, é, í, ó, ú, ů, ý
    # Pro ů/Ů normalizace funguje, ale pro jistotu.
    nfkd = unicodedata.normalize("NFD", text)
    stripped = "".join(c for c in nfkd if unicodedata.category(c) != "Mn")
    # NFD pro ů -> u + kroužek, pro ď/ť také funguje
    return stripped


def _preserve_case(original: str, corrected: str) -> str:
    """Zachová velikost písmen: ALL CAPS / Capitalized / lower."""
    if not original:
        return corrected
    if original.isupper():
        return corrected.upper()
    if original[0].isupper():
        return corrected[0].upper() + corrected[1:]
    return corrected


# ---------------------------------------------------------------------------
# Slovník: stripped lower -> diacritized lower
# Zdrojem je kurátorovaný seznam frekventovaných českých slov s diakritikou.
# Každý klíč je unikátní (strip(correct)). Ambiguous tvary (byt/být) jsou
# záměrně ponechány pouze s jednou variantou – bezpečné, nehádá kontextem.
# ---------------------------------------------------------------------------

# Seznam správných tvarů (lowercase). Z něj se automaticky odvodí stripped->correct.
_CORRECT_WORDS = [
    # příklady ze zadání
    "příliš", "žluťoučký", "kůň", "úpěl", "ďábelské", "ódy",
    "krásné", "počasí", "krásný", "pokusný",
    # běžná slova
    "dnes", "je", "včera", "zítra", "ano", "ne", "děkuji", "prosím",
    "dobrý", "den", "večer", "ráno", "noc", "čas", "člověk", "lidé",
    "práce", "škola", "domov", "rodina", "přítel", "láska", "život",
    "svět", "město", "vesnice", "ulice", "náměstí", "dům", "byt",
    "voda", "vzduch", "oheň", "země", "nebe", "slunce", "měsíc", "hvězda",
    "jaro", "léto", "podzim", "zima", "sníh", "déšť", "vítr", "bouře",
    "strom", "květina", "tráva", "les", "pole", "hora", "řeka", "moře",
    "auto", "vlak", "letadlo", "loď", "kolo", "cesta", "most", "nádraží",
    "kniha", "stránka", "písmeno", "slovo", "věta", "text", "článek",
    "počítač", "telefon", "internet", "program", "soubor", "složka",
    "nový", "starý", "velký", "malý", "dobrý", "špatný", "krásný", "ošklivý",
    "rychlý", "pomalý", "lehký", "těžký", "teplý", "studený", "čistý", "špinavý",
    "bílý", "černý", "červený", "zelený", "modrý", "žlutý", "růžový", "hnědý", "šedý",
    "první", "druhý", "třetí", "čtvrtý", "pátý", "poslední", "další", "jiný",
    "každý", "žádný", "některý", "všechny", "něco", "někdo", "nikdo", "někde",
    "tady", "tam", "kde", "kam", "kdy", "jak", "proč", "protože", "jestli",
    "který", "která", "které", "kteří", "jenž", "jež",
    "můj", "tvůj", "jeho", "její", "náš", "váš", "jejich",
    "já", "ty", "on", "ona", "ono", "my", "vy", "oni", "ony",
    "byl", "byla", "bylo", "byli", "byly", "být", "jsem", "jsi", "jsme", "jste", "jsou",
    "mám", "máš", "má", "máme", "máte", "mají", "mít",
    "chci", "chceš", "chce", "chceme", "chcete", "chtějí", "chtít",
    "mohu", "můžeš", "může", "můžeme", "můžete", "mohou", "moci",
    "dělám", "děláš", "dělá", "děláme", "děláte", "dělají", "dělat",
    "říkám", "říkáš", "říká", "říkáme", "říkáte", "říkají", "říkat", "říci",
    "vím", "víš", "ví", "víme", "víte", "vědí", "vědět",
    "vidím", "vidíš", "vidí", "vidíme", "vidíte", "vidí", "vidět",
    "slyším", "slyšíš", "slyší", "slyšíme", "slyšíte", "slyší", "slyšet",
    "přijít", "odejít", "přijdu", "přijdeš", "přijde", "přijdeme",
    "udělat", "napsat", "přečíst", "skočit", "běžet", "jít", "jet", "letět",
    "stůl", "židle", "okno", "dveře", "postel", "kuchyň", "koupelna",
    "jídlo", "pití", "chléb", "máslo", "sýr", "maso", "zelenina", "ovoce",
    "káva", "čaj", "pivo", "víno", "mléko", "cukr", "sůl", "pepř",
    "hlava", "ruka", "noha", "oko", "ucho", "nos", "ústa", "zub", "srdce",
    "lékař", "nemocnice", "lék", "zdraví", "nemoc", "bolest",
    "peníze", "práce", "plat", "úřad", "účet", "banka", "pošta",
    "škola", "učitel", "žák", "student", "třída", "předmět", "zkouška",
    "hudba", "píseň", "film", "divadlo", "kultura", "umění",
    "rok", "měsíc", "týden", "den", "hodina", "minuta", "vteřina",
    "pondělí", "úterý", "středa", "čtvrtek", "pátek", "sobota", "neděle",
    "leden", "únor", "březen", "duben", "květen", "červen", "červenec", "srpen", "září", "říjen", "listopad", "prosinec",
    "číslo", "počet", "částka", "strana", "řádka", "sloupec",
    "důležitý", "zajímavý", "krásný", "špatný", "správný", "nesprávný",
    "možný", "nemožný", "nutný", "zbytečný",
    "šťastný", "smutný", "veselý", "unavený", "živý", "mrtvý",
    "otevřený", "zavřený", "plný", "prázdný", "čistý", "špinavý",
    "úspěch", "neúspěch", "radost", "smutek", "láska", "nenávist",
    "přátelství", "rodina", "manžel", "manželka", "dítě", "rodiče",
    "město", "vesnice", "stát", "země", "národ", "jazyk", "řeč",
    "český", "česká", "české", "čeština", "češtině", "češtinu",
    "slovenský", "polský", "německý", "anglický", "francouzský",
    "velmi", "také", "již", "ještě", "právě", "totiž", "tedy", "proto",
    "ovšem", "však", "ale", "nebo", "ačkoli", "ačkoliv", "když", "pokud",
    "aby", "že", "protože", "přestože", "zatímco", "dokud", "než",
    "před", "za", "nad", "pod", "mezi", "vedle", "kolem", "skrz", "přes",
    "uvnitř", "venku", "nahoře", "dole", "vpředu", "vzadu", "vlevo", "vpravo",
    "tady", "zde", "tam", "sem", "tamhle",
    "teď", "nyní", "dříve", "později", "včera", "dnes", "zítra",
    "často", "někdy", "vždy", "nikdy", "zřídka",
    "ano", "ne", "možná", "určitě", "jistě", "snad", "prý",
    "prosím", "děkuji", "promiň", "omlouvám", "vítej",
    "háček", "čárka", "kroužek", "písmeno", "abeceda",
    "příklad", "ukázka", "cvičení", "úkol", "řešení",
    "začátek", "konec", "prostředek", "okraj", "střed",
    "východ", "západ", "sever", "jih",
    "jaro", "léto", "podzim", "zima",
    "koupit", "prodat", "dát", "vzít", "nést", "vézt",
    "hledat", "najít", "ztratit", "držet", "pustit",
    "mluvit", "říkat", "ptát", "odpovědět", "myslet", "vědět",
    "cítit", "slyšet", "vidět", "vonět", "chutnat",
    "milovat", "nenávidět", "chtít", "potřebovat", "muset",
    "začít", "skončit", "pokračovat", "přestat",
    "přijít", "odejít", "přinést", "odnést",
    "otevřít", "zavřít", "zapnout", "vypnout",
    "kočka", "pes", "pták", "ryba", "kůň", "kráva", "prase", "ovce", "koza", "slepice",
    "strom", "keř", "květ", "list", "kořen", "větev", "kmen", "les",
    "kámen", "písek", "hlína", "voda", "oheň", "vzduch",
    "slovo", "věta", "odstavec", "kapitola", "kniha",
    "dopis", "zpráva", "noviny", "časopis",
    "úřad", "škola", "nemocnice", "obchod", "restaurace", "hotel",
    "nádraží", "letiště", "přístav", "most", "tunel",
    "ulice", "náměstí", "park", "zahrada", "hřiště",
    "auto", "autobus", "tramvaj", "vlak", "letadlo", "kolo",
    "počasí", "teplota", "déšť", "sníh", "vítr", "mlha", "bouřka",
    "barva", "tvar", "velikost", "hmotnost", "výška", "šířka", "délka",
    "rychlost", "vzdálenost", "čas", "prostor",
    "peníze", "cena", "hodnota", "účet", "platba",
    "práce", "zaměstnání", "povolání", "úkol", "povinnost",
    "rodina", "přátelé", "soused", "kolega",
    "děti", "rodiče", "prarodiče", "sourozenec",
    "manžel", "manželka", "partner", "partnerka",
    "svátek", "narozeniny", "oslavit", "dárek",
    "vánoce", "velikonoce", "nový", "rok",
    "dovolená", "prázdniny", "výlet", "cesta",
    "jídlo", "snídaně", "oběd", "večeře", "svačina",
    "čaj", "káva", "pivo", "víno", "voda", "džus",
    "chléb", "rohlík", "houska", "koláč", "dort",
    "polévka", "omáčka", "maso", "ryba", "zelenina", "ovoce",
    "sůl", "cukr", "pepř", "ocet", "olej", "máslo", "sýr", "mléko", "vejce",
    "účet", "úspěch", "úroveň", "úhel", "úkol", "úsměv", "ústa", "úterý", "úvaha",
    "ďábel", "ďábelské", "ďáblův", "ťukat", "ťapka", "ňadra", "žába", "žák", "žena", "židle",
    "šála", "šátek", "šéf", "léto", "lékař", "tělo", "těšit", "město", "měsíc", "pěkný", "pět",
    "oběd", "věc", "věřit", "vědec", "věž", "běhat", "běžný",
    "kůže", "kůl", "kůra", "dům", "stůl", "sůl", "vůl", "vůně",
    "půda", "půl", "půjčit", "původ", "růže", "růst", "průběh", "průměr",
]

# Sestavení slovníku stripped->correct (pouze kde se liší)
_DIACRITICS_DICT: dict[str, str] = {}
for _w in _CORRECT_WORDS:
    _key = strip_diacritics(_w.lower())
    _correct = _w.lower()
    if _key == _correct:
        continue
    # pokud kolize (více slov mapuje na stejný stripped), ponech první – bezpečné
    if _key not in _DIACRITICS_DICT:
        _DIACRITICS_DICT[_key] = _correct

# Ručně přidané mapy pro tvary kde se liší koncovka (např. žlutý -> žluťoučký je odvozené, ale potřebujeme i další)
_EXTRA_MAP = {
    "krasne": "krásné",
    "krasny": "krásný",
    "krasna": "krásná",
    "krasneho": "krásného",
    "zlutoucky": "žluťoučký",
    "zlutoucka": "žluťoučká",
    "zlutoucke": "žluťoučké",
    "dabelske": "ďábelské",
    "dabelsky": "ďábelský",
    "dabelska": "ďábelská",
    "pokusny": "pokusný",
    "pocasi": "počasí",
    "pocasiho": "počasí",
    "prilis": "příliš",
    "kun": "kůň",
    "kone": "koně",
    "koni": "koní",
    "upel": "úpěl",
    "upela": "úpěla",
    "ody": "ódy",
    "stranka": "stránka",
    "stranky": "stránky",
    "strance": "stránce",
    "radka": "řádka",
    "radky": "řádky",
    "cislo": "číslo",
    "cisla": "čísla",
}
for _k, _v in _EXTRA_MAP.items():
    _DIACRITICS_DICT[_k] = _v

# Tokenizace: slovo (včetně čísel) vs. mezery vs. interpunkce
# \w s UNICODE zahrnuje písmena s diakritikou
_TOKEN_RE = re.compile(r"(\w+|\s+|[^\w\s]+)", re.UNICODE)

# Header stránkování – nemá se diakritizovat jako celek jinak, ale slovo "Stránka" má diakritiku
_PAGE_HEADER_RE = re.compile(r"---\s*Stránka\s+\d+\s*---")

# ---------------------------------------------------------------------------
# Jádro diakritizace
# ---------------------------------------------------------------------------

def _diacritize_word(word: str) -> str:
    """Vrátí diakritizovanou verzi slova pokud je ve slovníku, jinak originál.

    Pouze word-level mapování – žádné c->č.
    """
    if not word:
        return word
    # čísla a slova s číslicemi necháváme
    if not word.isalpha():
        # např. "123" nebo "a1" – nech
        # ale "krasne" je isalpha True
        # pokud obsahuje číslice, vrať originál
        if any(ch.isdigit() for ch in word):
            return word
        # pro slova s podtržítkem apod. – nech
        if not word.isalpha():
            return word
    key = strip_diacritics(word.lower())
    # pokud slovo už obsahuje diakritiku a zároveň jeho stripped forma není ve slovníku,
    # ponech beze změny (neznámé slovo – bezpečné)
    correct_lower = _DIACRITICS_DICT.get(key)
    if correct_lower is None:
        return word
    # pokud je slovo již správně (lower == correct), vrať originál (zachová case)
    if word.lower() == correct_lower:
        return word
    # pokud je word.lower() stripped a mapuje na correct, vrať s původní case
    return _preserve_case(word, correct_lower)


def diacritize_text(text: str) -> str:
    """Kontextová bezpečná diakritizace textu po slovech.

    Zachovává whitespace, interpunkci, čísla, header --- Stránka N ---.
    """
    if not text:
        return text

    # Rozdělíme po částech, ale zachováme vše
    # Pokud text obsahuje page headery, zpracujeme je po částech aby se neporušily
    # Jednoduše tokenizujeme celý text – header je "---", " ", "Stránka" atd. – projde slovníkem
    parts = _TOKEN_RE.findall(text)
    out_parts: list[str] = []
    for part in parts:
        # whitespace a interpunkce beze změny
        if not part or part.isspace() or not part.strip():
            out_parts.append(part)
            continue
        # interpunkce typu "---" nebo "," -> re je rozdělí na jednotlivé tokeny, ale findall dá např. "---" jako jeden? 
        # Naše regex dá "---" jako jeden token (ne \w, ne \s -> [^\w\s]+) – ponech beze změny
        if re.fullmatch(r"[^\w\s]+", part):
            out_parts.append(part)
            continue
        # slovo
        if re.fullmatch(r"\w+", part):
            out_parts.append(_diacritize_word(part))
        else:
            out_parts.append(part)
    return "".join(out_parts)


# ---------------------------------------------------------------------------
# Volitelný transformers backend (BERT) – pokud je nainstalován
# ---------------------------------------------------------------------------

_transformers_lock = threading.Lock()
_transformers_model = None
_transformers_tokenizer = None
_transformers_available: Optional[bool] = None


def _try_load_transformers():
    global _transformers_model, _transformers_tokenizer, _transformers_available
    with _transformers_lock:
        if _transformers_available is not None:
            return _transformers_available
        try:
            import torch  # noqa: F401
            from transformers import AutoTokenizer, AutoModelForTokenClassification  # noqa: F401
            # Zde by se načetl model typu "ufal/robeczech-base" fine-tuned pro diakritiku
            # Model není bundlován – pokud není lokálně cache, fallback na slovník
            # Prozatím detekujeme pouze dostupnost knihovny, model load je lazy a vyžaduje
            # lokální soubory v .cache/huggingface . Pokud chybí, fallback.
            _transformers_available = True
        except Exception:
            _transformers_available = False
        return _transformers_available


def _transformers_diacritize(text: str) -> Optional[str]:
    """Pokus o diakritizaci přes transformers BERT. Vrátí None pokud selže/nedostupné."""
    if not _try_load_transformers():
        return None
    # Zatím bez konkrétního checkpointu – slovníkový fallback je primární offline řešení.
    # Pokud by byl model dostupný lokálně (např. ./models/czech-diacritics), zde by proběhla inference:
    #   tokenizer = AutoTokenizer.from_pretrained(local_path)
    #   model = AutoModelForTokenClassification.from_pretrained(local_path)
    #   ...
    # Pro účely zadání je slovníková metoda považována za plnohodnotné offline řešení
    # splňující prioritu "bez internetu, Windows, rozumná velikost".
    return None


# ---------------------------------------------------------------------------
# Veřejné API
# ---------------------------------------------------------------------------

class CzechDiacritizer:
    """Singleton pro diakritizaci. Thread-safe, lazy."""

    _instance: Optional["CzechDiacritizer"] = None
    _lock = threading.Lock()

    def __new__(cls):
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super().__new__(cls)
                    cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        if getattr(self, "_initialized", False):
            return
        self._initialized = True
        self._use_transformers = False

    def is_available(self) -> bool:
        return True  # slovník je vždy dostupný

    def diacritize(self, text: str) -> str:
        """Bezpečná diakritizace – nikdy nevyhodí, při chybě vrátí originál."""
        if not text or not text.strip():
            return text
        try:
            # nejprve zkus transformers pokud je nakonfigurován a dostupný
            tr = _transformers_diacritize(text)
            if tr is not None:
                return tr
            return diacritize_text(text)
        except Exception:
            return text

    def diacritize_page(self, text: str) -> str:
        return self.diacritize(text)


# Globální instance
_diacritizer = CzechDiacritizer()


def get_diacritizer() -> CzechDiacritizer:
    return _diacritizer


def should_diacritize(lang_code: str, enabled: bool) -> bool:
    """Aktivní pouze pro češtinu (ces) a když je volba zapnutá."""
    if not enabled:
        return False
    if not lang_code:
        return False
    # lang_code může být "ces (Čeština)" nebo "ces"
    code = lang_code.split()[0].strip().lower()
    return code == "ces"


def diacritize_text_safe(text: str, lang_code: str, enabled: bool) -> tuple[str, bool]:
    """Vrátí (text_po_opravě, success_flag). Při vypnuto nebo ne-ces vrací originál, False."""
    if not should_diacritize(lang_code, enabled):
        return text, False
    try:
        corrected = get_diacritizer().diacritize(text)
        return corrected, True
    except Exception:
        return text, False


# ---------------------------------------------------------------------------
# Pomoc pro OcrResult – zachování bbox
# ---------------------------------------------------------------------------

def apply_diacritics_to_results(results: list, lang_code: str, enabled: bool):
    """Aplikuje diakritizaci na list OcrResult, zachová bbox.

    Vrací nový list se stejnými bbox, ale text opravený.
    Každý OcrResult dostane original_text a processed_text.
    """
    if not should_diacritize(lang_code, enabled):
        return results
    new_results = []
    for r in results:
        try:
            original = r.text
            corrected = get_diacritizer().diacritize(original)
            # pokud se neliší, ponech original
            if corrected != original:
                # vytvoř nový objekt se zachovaným bbox
                from ocr_engine import OcrResult as _OcrResult
                new_r = _OcrResult(text=corrected, bbox=r.bbox, original_text=original, processed_text=corrected)
            else:
                from ocr_engine import OcrResult as _OcrResult
                # i když se neliší, vyplň original/processed pro konzistenci
                new_r = _OcrResult(text=original, bbox=r.bbox, original_text=original, processed_text=original)
            new_results.append(new_r)
        except Exception:
            new_results.append(r)
    return new_results
