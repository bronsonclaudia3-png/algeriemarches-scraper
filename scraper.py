"""
AlgerieMarches Scraper v4 — Dual Scrape (Appels d'Offres & Avis d'Attributions)
Automated daily sync to Google Sheets & local Downloads Excel file.

Targets:
  1. "APPEL D'OFFRE 2026" (11 columns)
  2. "AVIS D'ATTRIBUTIONS 2026" (15 columns)
Domain (TAPIDOR):
  GAZON — turf for stadiums, sports grounds, playgrounds, school courtyards.
  TAPIS — carpets & moquette for mosques and buildings.
"""

import argparse
import base64
import csv
import json
import logging
import os
import re
import shutil
import sys
import time
import unicodedata
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import unquote

import gspread
import requests
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload
from gspread.utils import rowcol_to_a1

# ── Paths ────────────────────────────────────────────────────────────────────
SCRIPT_DIR = Path(__file__).resolve().parent
STATE_DIR = SCRIPT_DIR / "state"
DATA_DIR = SCRIPT_DIR / "data"
SCANS_DIR = DATA_DIR / "scans"
COOKIE_FILE = STATE_DIR / "cookies.json"
LOG_FILE = SCRIPT_DIR / "scraper.log"

STATE_DIR.mkdir(exist_ok=True)
DATA_DIR.mkdir(exist_ok=True)
SCANS_DIR.mkdir(parents=True, exist_ok=True)

# ── Logging ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(
            open(sys.stdout.fileno(), mode="w", encoding="utf-8", closefd=False)
        ),
    ],
)
log = logging.getLogger("am-scraper")

# ── Config ───────────────────────────────────────────────────────────────────
load_dotenv(SCRIPT_DIR / ".env")

EMAIL = os.getenv("AM_EMAIL", "direction@tapidor.com")
PASSWORD = os.getenv("AM_PASSWORD", "")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
NVIDIA_API_KEY = os.getenv("NVIDIA_API_KEY", "").strip()
GOOGLE_DRIVE_FOLDER_ID = os.getenv("GOOGLE_DRIVE_FOLDER_ID", "").strip()

BASE_URL = "https://algeriemarches.com"
API_BASE = "https://api.algeriemarches.com/api"
API_PROXY = f"{BASE_URL}/api/proxy/ads"

CSRF_URL = f"{BASE_URL}/api/auth/csrf"
LOGIN_URL = f"{BASE_URL}/api/auth/callback/credentials"
SESSIONS_URL = f"{BASE_URL}/api/auth/active-sessions"
DISCONNECT_URL = f"{BASE_URL}/api/proxy/auth/disconnect-session"

PAGE_SIZE = 20
MAX_PAGES = 15
# A 7-day window is ~950-1000 notices per type, so the catch-up cap must sit well above 50 pages
MAX_PAGES_CATCHUP = 120
LISTING_RETRIES = 3
LISTING_RETRY_DELAY = 5  # seconds, multiplied by the attempt number

# Per-notice verdicts so a notice is OCR'd once, not on every daily run of its 7-day window.
# Bump VERDICTS_VERSION whenever the scope rules or the prompt change, to re-check old verdicts.
VERDICTS_FILE = STATE_DIR / "verdicts.json"
VERDICTS_VERSION = 2
VERDICTS_MAX_AGE_DAYS = 60
MAX_SCAN_PAGES = 3

# ── TAPIDOR scope keywords ───────────────────────────────────────────────────
# TAPIDOR works on two product families:
#   GAZON — turf: gazon / pelouse / engazonnement / عشب / تعشيب
#   TAPIS — carpets & moquette for mosques and buildings: tapis / moquette / زرابي / سجاد / موكيت
# The scope regexes are applied to normalize_text() output (lowercase, no accents, no Arabic
# diacritics, unified alef/ta marbuta) but stay accent-tolerant so they also work on raw text.

DEFAULT_KEYWORDS = (
    r"\b(?:gazons?|gazonn\w*|pelouses?|engazonn\w*|stades?|football|sport|"
    r"terrains?\s+(?:de\s+)?(?:sports?|sportifs?|football|foot|jeux?|proximit[eé]|omnisports?|multisports?|combin[eé]s?)|"
    r"aires?\s+de\s+jeux?|tapis|moquettes?|moquettage)\b"
    r"|عشب|تعشيب|نجيل|ملعب|ملاعب|زراب|زربي|سجاد|موكيت"
    r"|\b(?:[eé]quipements?|ameublements?|fournitures?)\b.{0,60}\bmosqu[eé]es?\b"
    r"|(?:تجهيز|تاثيث|تأثيث|فرش).{0,40}(?:مسجد|مساجد)"
)
KEYWORDS_ENV = os.getenv("AM_KEYWORDS", "").strip()
ACTIVE_KEYWORDS = KEYWORDS_ENV if KEYWORDS_ENV else DEFAULT_KEYWORDS
KEYWORD_FILTER = re.compile(ACTIVE_KEYWORDS, re.IGNORECASE)

# Arabic words take attached prefixes (ال، بال، لل، و), so Arabic terms are matched without \b.
GAZON_EXPLICIT_KEYWORDS = re.compile(
    r"\b(?:gazons?|gazonn[eé]e?s?|pelouses?|engazonn\w*)\b"
    r"|عشب|تعشيب|(?<![اإ])نجيل|تكسي[ةه]\s+[اإ]صطناعي[ةه]",
    re.IGNORECASE
)

# Mentions of gazon that are equipment or upkeep, not a turf project
GAZON_OUT_OF_SCOPE = re.compile(
    r"\b(?:tondeuses?|tracteurs?[\s-]tondeuses?)\s+(?:[aà]|de|pour)\s+gazons?\b"
    r"|\bcoupe[\s-]?gazons?\b|\bd[eé]broussailleuses?\b"
    r"|\btontes?\s+(?:du\s+|de\s+la\s+|des\s+)?(?:gazons?|pelouses?)\b",
    re.IGNORECASE
)

TAPIS_EXPLICIT_KEYWORDS = re.compile(
    r"\b(?:tapis|moquettes?|moquettage)\b|زراب[يى]|زربي[ةه]|سجاد|موكيت",
    re.IGNORECASE
)

# "tapis" that is not a carpet: treadmills, conveyors, asphalt wearing courses, sports mats
TAPIS_OUT_OF_SCOPE = re.compile(
    r"\btapis\s+(?:roulants?|de\s+course|convoyeurs?|transporteurs?|de\s+souris|"
    r"d['’\s]?enrob\w*|en\s+enrob\w*|bitum\w*|en\s+b\.?\s?b\b|en\s+b[eé]ton\s+bitum\w*|"
    r"de\s+(?:judo|lutte|karat[eé]|taekwondo|gym\w*|yoga|fitness|sol\s+sporti\w*|r[eé]ception)|"
    r"anti[\s-]?d[eé]rap\w*|[aà]\s+bagages|de\s+bain|d['’\s]?[eé]vacuation|d['’\s]?usure)\b"
    r"|\bconvoyeurs?\s+[aà]\s+tapis\b|\btatamis?\b|سير\s+ناقل|بساط\s+متحرك",
    re.IGNORECASE
)

# A "tapis" next to asphalt vocabulary is a road wearing course
ROAD_SURFACE_CONTEXT = re.compile(
    r"\b(?:enrob\w*|bitum\w*|chauss[eé]es?|couche\s+de\s+roulement)\b",
    re.IGNORECASE
)

NON_GAZON_EXCLUSIONS = re.compile(
    r"\b(?:cour\s+de\s+justice|cours\s+de\s+justice|tribunal|tribunaux|مجلس\s+قضاء|محكمة|"
    r"chauffage|climatisation|plomberie|electricite|électricité|éclairage\s+public|eclairage\s+public|"
    r"mobilier|bureau|fourniture\s+de\s+bureau|fournitures\s+de\s+bureau|عتاد\s+مكتبي|أثاث|اثاث|طباعة|"
    r"assiette\s+de\s+terrain|assiette\s+fonciere|assiette\s+foncière|logement|logements|aadl|lpp|lpa|pos|viabilisation|geotechnique|géotechnique|"
    r"assainissement|assainisement|aep|drainage|cantine|salles?\s+omnisports?|salles?\s+de\s+sports?|piscine|"
    r"véhicule|véhicules|vehicule|vehicules|trémie|tremie|gradin|gradins|"
    r"glissement|glissements|route\s+nationale|chemin\s+communal|chemins\s+communaux|ouvrage\s+d'art|pont|ponts|voie|voirie|trottoir|trottoirs|"
    r"piste\s+d['’]atterrissage|a[eé]roport|a[eé]rodrome|aviation|bombardier|atterrissage|avion|vol|"
    r"incendie|incendies|for[eê]t|for[eê]ts|boisement|reboisement|"
    r"barrage|retenue\s+collinaire|forage|ch[aâ]teau\s+d['’]eau|hydraulique|"
    r"port|ports|maritime|p[eê]che|"
    r"routes?|autoroutes?|d[eé]viation|tranch[eé]e|rev[eê]tement\s+en\s+enrob[eé]|enrob[eé]|enrob[eé]s|bicouche|tricouche|"
    r"bitum\w*|chauss[eé]es?|couche\s+de\s+roulement|mat[eé]riels?\s+agricoles?|"
    r"psychop[eé]dagogique|handicap[eé]|handicap[eé]s)\b",
    re.IGNORECASE
)

# Exclusions that often appear inside genuine turf projects (stadium drainage, stands, lighting,
# sports complexes). They no longer discard a notice outright; the scan is checked instead.
SOFT_EXCLUSIONS = re.compile(
    r"\b(?:drainage|gradins?|[eé]clairage\s+public|piscines?|salles?\s+omnisports?|"
    r"voiries?|trottoirs?|assainissement|assainisement)\b",
    re.IGNORECASE
)

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
)

# ── Classification & Parsing Helpers ─────────────────────────────────────────

def strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")

def normalize_text(s) -> str:
    """Lowercase, accent-free, diacritic-free text for keyword matching (French & Arabic)."""
    if not s:
        return ""
    s = str(s)
    # AlgerieMarches titles carry a broken "à" (U+FFFD + NBSP)
    s = s.replace("� ", "à").replace("�", "")
    s = strip_accents(s).lower()
    s = s.replace("ـ", "")  # Arabic tatweel
    s = re.sub("[إأآٱ]", "ا", s).replace("ى", "ي").replace("ة", "ه")
    s = s.replace("’", "'").replace("‘", "'")
    return re.sub(r"[\s   ]+", " ", s).strip()

def detect_scope(text: str) -> str | None:
    """Return "GAZON" or "TAPIS" when the text explicitly names a TAPIDOR product, else None."""
    t = normalize_text(text)
    if GAZON_EXPLICIT_KEYWORDS.search(GAZON_OUT_OF_SCOPE.sub(" ", t)):
        return "GAZON"
    t_tapis = TAPIS_OUT_OF_SCOPE.sub(" ", t)
    if TAPIS_EXPLICIT_KEYWORDS.search(t_tapis) and not ROAD_SURFACE_CONTEXT.search(t_tapis):
        return "TAPIS"
    return None

def prefilter_decision(text: str) -> tuple[str, str]:
    """
    Decide what to do with a listing before fetching its detail page:
      ("explicit", scope) — title/description names gazon or tapis: keep it.
      ("ocr", "")         — sports/mosque wording without an explicit product: read the scan.
      ("skip", reason)    — unrelated notice.
    """
    scope = detect_scope(text)
    if scope:
        return "explicit", scope
    t = normalize_text(text)
    t_clean = GAZON_OUT_OF_SCOPE.sub(" ", TAPIS_OUT_OF_SCOPE.sub(" ", t))
    # Custom AM_KEYWORDS may be written with accents, so also try them on the raw text
    if not (KEYWORD_FILTER.search(t_clean) or (KEYWORDS_ENV and KEYWORD_FILTER.search(str(text or "")))):
        return "skip", "no keyword"
    hard = NON_GAZON_EXCLUSIONS.search(SOFT_EXCLUSIONS.sub(" ", t_clean))
    if hard:
        return "skip", f"excluded ({hard.group(0)})"
    return "ocr", ""

def classify_action(titre: str) -> str:
    """Classify the action verb for Col 3 (TITRE D'APPEL D'OFFRE / TITRE D'AVIS D'ATTRIBUTION)."""
    t_norm = strip_accents(titre).upper()

    patterns = [
        ("AMÉNAGEMENT ET REVÊTEMENT", r"\bAMENAGEMENT\b.*\bREVETEMENT\b|\bREVETEMENT\b.*\bAMENAGEMENT\b"),
        ("RÉALISATION ET REVÊTEMENT", r"\bREALISATION\b.*\bREVETEMENT\b|\bREVETEMENT\b.*\bREALISATION\b"),
        ("RÉALISATION ET RÉHABILITATION", r"\bREALISATION\b.*\bREHABILITATION\b|\bREALISATION\b.*\bRH[EÉ]ABILITATION\b"),
        ("RÉALISATION ET AMÉNAGEMENT", r"\bREALISATION\b.*\bAMENAGEMENT\b|\bAMENAGEMENT\b.*\bREALISATION\b"),
        ("AMÉNAGEMENT ET RÉHABILITATION", r"\bAMENAGEMENT\b.*\bREHABILITATION\b|\bREHABILITATION\b.*\bAMENAGEMENT\b"),
        ("RHÉABILITATION ET REVÊTEMENT", r"\bRH?EABILITATION\b.*\bREVETEMENT\b|\bREVETEMENT\b.*\bRH?EABILITATION\b"),
        ("AMÉNAGEMENT ET COUVERTURE", r"\bAMENAGEMENT\b.*\bCOUVERTURE\b|\bCOUVERTURE\b.*\bAMENAGEMENT\b"),
        ("TRAVAUX D'ENGAZONNEMENT", r"\bTRAVAUX D'?ENGAZONNEMENT\b|\bENGAZONNEMENT\b"),
        ("FOURNITURE ET POSE", r"\bFOURNITURE ET POSE\b|\bFOURNITURE\b.*\bPOSE\b"),
        ("FOURNITURE ET INSTALLATION", r"\bFOURNITURE ET INSTALLATION\b"),
        ("REMISE À NIVEAU", r"\bREMISE A NIVEAU\b"),
        ("RÉALISATION", r"\bREALISATION\b"),
        ("AMÉNAGEMENT", r"\bAMENAGEMENT\b"),
        ("REVÊTEMENT", r"\bREVETEMENT\b|\bREVENTEMENT\b"),
        ("RÉHABILITATION", r"\bREHABILITATION\b|\bRH[EÉ]ABILITATION\b|\bREHABITATION\b"),
        ("FOURNITURE", r"\bFOURNITURE\b"),
        ("RÉNOVATION", r"\bRENOVATION\b"),
        ("ACHÈVEMENT", r"\bACHEVEMENT\b"),
        ("CONSTRUCTION", r"\bCONSTRUCTION\b"),
        ("RÉFECTION", r"\bREFECTION\b"),
        ("COUVERTURE", r"\bCOUVERTURE\b"),
        ("ÉTUDE ET SUIVI", r"\bETUDE ET SUIVI\b|\bETUDE\b.*\bSUIVI\b"),
        ("SUIVI ET RÉALISATION", r"\bSUIVI\b.*\bREALISATION\b"),
        ("SUIVI ET AMÉNAGEMENT", r"\bSUIVI\b.*\bAMENAGEMENT\b"),
        ("TRAVAUX", r"\bTRAVAUX\b"),
    ]
    for action_label, pattern in patterns:
        if re.search(pattern, t_norm):
            return action_label
    return "TRAVAUX"

def classify_type_projet(titre: str) -> str:
    """Classify the facility / project type for Col 5 (TYPE DE PROJET)."""
    t_norm = strip_accents(titre).upper()

    if detect_scope(titre) == "TAPIS":
        t_low = normalize_text(titre)
        if re.search(r"\bmosqu|مسجد|مساجد", t_low):
            return "TAPIS MOSQUÉE"
        if re.search(r"\bmoquett|موكيت", t_low):
            return "MOQUETTE"
        return "TAPIS"

    facility_patterns = [
        ("STADE DE PROXIMITÉ", r"\bSTADES?\s+(?:DE\s+)?PROXIMIT[EÉ]\b"),
        ("STADE DE FOOTBALL", r"\bSTADES?\s+DE\s+FOOTBALL\b|\bSTADES?\s+DE\s+FOOT\b"),
        ("STADE COMMUNAL", r"\bSTADES?\s+COMMUNAL(?:ES?|S)?\b|\bSTADES?\s+COMMUNAUX\b"),
        ("STADE MUNICIPAL", r"\bSTADES?\s+MUNICIPAL(?:ES?|S)?\b|\bSTADES?\s+MUNICIPAUX\b"),
        ("STADE MATICO", r"\bSTADES?\s+MATICO\b"),
        ("STADE SPORTIF", r"\bSTADES?\s+SPORTIFS?\b"),
        ("STADE", r"\bSTADES?\b"),
        ("TERRAIN DE FOOTBALL", r"\bTERRAINS?\s+DE\s+FOOTBALL\b|\bTERRAINS?\s+DE\s+FOOT\b"),
        ("TERRAIN SPORTIF DE PROXIMITÉ", r"\bTERRAINS?\s+SPORTIFS?\s+DE\s+PROXIMIT[EÉ]\b"),
        ("TERRAIN DE PROXIMITÉ", r"\bTERRAINS?\s+DE\s+PROXIMIT[EÉ]\b"),
        ("TERRAIN DE SPORT", r"\bTERRAINS?\s+DE\s+SPORTS?\b|\bTERRAINS?\s+SPORTIFS?\b"),
        ("TERRAIN DE JEU", r"\bTERRAINS?\s+DE\s+JEUX?\b"),
        ("TERRAIN GAZONNÉ", r"\bTERRAINS?\s+GAZONN[EÉ]S?\b|\bTERRAINS?\s+GAZONS?\b"),
        ("TERRAIN", r"\bTERRAINS?\b"),
        ("AIRE DE JEUX", r"\bAIRES?\s+DE\s+JEUX?\b|\bESPACES?\s+DE\s+JEUX?\b|\bAIR\s+DE\s+JEUX?\b"),
        ("MATICO", r"\bMATICO\b|\bMATIQUO\b"),
        ("COUR ECOLE PRIMAIRE", r"\bCOURS?\s+(?:D'?)?ECOLES?\s+PRIMAIRES?\b|\bECOLES?\s+PRIMAIRES?\b|\bCOURS?\s+(?:D'?)?ECOLES?\b"),
        ("COUR", r"\bCOURS?\b"),
        ("PISTE D'ATHLETISME", r"\bPISTES?\s+(?:D'?)?ATHLETISME\b|\bATHLETISME\b"),
        ("SALLE DE SPORT", r"\bSALLES?\s+MULTI[\s\-]SPORTS?\b|\bSALLES?\s+DE\s+SPORTS?\b|\bCOMPLEXES?\s+SPORTIFS?\b|\bSALLE\s+OMNISPORTS?\b"),
        ("GAZON SYNTHÉTIQUE", r"\bGAZONS?\s+SYNTH[EÉ]TIQUES?\b|\bPELOUSES?\s+SYNTH[EÉ]TIQUES?\b|\bGAZONS?\b"),
    ]
    for facility_label, pattern in facility_patterns:
        if re.search(pattern, t_norm):
            return facility_label
    return "INFRASTRUCTURE SPORTIVE"

COMMUNE_NAME_STOP_WORDS = {
    "WILAYA", "DAIRA", "LOT", "LOTS", "POUR", "AU", "AUX", "A", "EN", "ET", "DANS", "SUR", "AVEC",
    "PAR", "SIS", "SISE", "SITUE", "SITUEE", "LE", "LA", "LES", "DU", "DES", "DE", "D", "PROJET",
    "REALISATION", "AMENAGEMENT", "REHABILITATION", "TRAVAUX", "FOURNITURE", "ETUDE", "SUIVI",
    "ALIMENTATION", "QUARTIER", "LIEU", "CHEF",
}

def parse_commune(titre: str, annonceur: str) -> str:
    """Commune name after 'Commune de / APC de' in the annonceur, then in the title; '/' if absent."""
    for source in (annonceur, titre):
        u = strip_accents(str(source or "")).upper().replace("’", "'").replace("�", " ")
        m = re.search(
            r"\b(?:COMMUNE|APC|A\.P\.C)\b\s*(?:D'|DE\s+L'|DE\s+|DES\s+|DU\s+)?\s*([A-Z][A-Z'\- ]*)", u
        )
        if not m:
            continue
        words = []
        for w in m.group(1).split():
            if w.strip("'-") in COMMUNE_NAME_STOP_WORDS:
                break
            words.append(w)
            if len(words) == 5:
                break
        name = " ".join(words).strip(" -'")
        if len(name) >= 3:
            return name
    return "/"

def clean_annonceur(ann_raw: str) -> str:
    if not ann_raw or not str(ann_raw).strip():
        return "/"
    ann_raw = str(ann_raw)
    u = strip_accents(ann_raw).upper().strip().replace("’", "'")
    if re.match(r"(?:COMMUNE|APC|A\.P\.C|ASSEMBLEE POPULAIRE COMMUNALE)\b", u):
        return "COMMUNE"
    if re.search(r"\bDJS\b|JEUNESSE ET (?:DES )?SPORTS?\b", u):
        return "DJS DE LA WILAYA"
    if re.search(r"\bD\.?E\.?P\b|EQUIPEMENTS? PUBLICS?\b", u):
        return "DEP DE LA WILAYA"
    if re.search(r"ADMINISTRATION LOCALE|\bD\.?A\.?L\b", u):
        return "DIRECTION DE L'ADMINISTRATION LOCALE DE LA WILAYA"
    if re.search(r"\bEDUCATION\b", u):
        return "DIRECTION DE L'EDUCATION DE LA WILAYA"
    if "FORMATION ET DE L'ENSEIGNEMENT" in u:
        return "DIRECTION DE LA FORMATION ET DE L'ENSEIGNEMENT PROFESSIONNELS"
    if re.search(r"\b(?:COMMUNE|APC|A\.P\.C)\b", u):
        return "COMMUNE"
    return ann_raw[:50].strip()

# A money token: grouped thousands (space, dot or comma) with optional 1-2 decimals, or a plain
# 4+ digit number. Look-arounds stop it from swallowing parts of dates (12/08/2026, 15.06.2026).
_AMOUNT_RE = re.compile(
    r"(?<![\d.,/-])(\d{1,3}(?:[ .,]\d{3})+(?:[.,]\d{1,2})?|\d{4,}(?:[.,]\d{1,2})?)(?![\d/]|[.,]\d)"
)
_AMOUNT_PARTS_RE = re.compile(r"^(\d+(?:[.,]\d{3})*)(?:[.,](\d{1,2}))?$")
_CURRENCY_RE = re.compile(r"^\s*(?:DA\b|DZD\b|D\.A\b|دج|د\.ج|DINARS?\b)", re.IGNORECASE)
_TTC_LABEL = r"(?:TTC|T\.T\.C|TOUTES\s+TAXES\s+COMPRISES)"
_TTC_AFTER_RE = re.compile(r"\s*(?:DA|DZD|D\.A|دج|د\.ج|DINARS?)?\.?\s*\(?\s*" + _TTC_LABEL, re.IGNORECASE)
_TTC_BEFORE_RE = re.compile(_TTC_LABEL + r"\s*\)?\s*(?::|=|EST\s+DE|SOIT|DE)?\s*$", re.IGNORECASE)

def _amount_to_number(token: str):
    m = _AMOUNT_PARTS_RE.match(token.replace(" ", ""))
    if not m:
        return None
    value = float(re.sub(r"[.,]", "", m.group(1)) + ("." + m.group(2) if m.group(2) else ""))
    return int(value) if value.is_integer() else value

def extract_amount(text):
    """
    Pick ONE amount out of free text without merging neighbouring numbers or dates.
    Prefers amounts followed by a currency (DA/DZD/دج); among several, the one marked TTC.
    Returns int/float, or None when no plausible amount (>= 1000 DA) is present.
    """
    if text is None:
        return None
    if isinstance(text, (int, float)) and not isinstance(text, bool):
        return int(text) if float(text).is_integer() else float(text)
    s = re.sub(r"[   ]", " ", str(text))
    candidates = []
    for m in _AMOUNT_RE.finditer(s):
        val = _amount_to_number(m.group(1))
        if val is None or val < 1000:
            continue
        candidates.append((m.start(), m.end(), val, bool(_CURRENCY_RE.match(s[m.end():]))))
    if not candidates:
        return None
    tagged = [c for c in candidates if c[3]]
    if not tagged:
        # Untagged numbers: drop bare years when something else is available
        non_years = [c for c in candidates if not (isinstance(c[2], int) and 1990 <= c[2] <= 2100)]
        tagged = non_years or candidates
    if len(tagged) > 1:
        for start, end, val, _ in tagged:
            if _TTC_AFTER_RE.match(s[end:]) or _TTC_BEFORE_RE.search(s[max(0, start - 40):start]):
                return val
    return tagged[0][2]

def parse_budget(montant_str):
    if montant_str is None or (isinstance(montant_str, str) and montant_str.strip() in ("", "/")):
        return "/"
    val = extract_amount(montant_str)
    return val if val is not None else "/"

_DELAI_RE = re.compile(
    r"(\d+)\s*\)?\s*(?:\([^)\d]{1,20}\)\s*)?(jours?|j\b|mois|semaines?|يوم|ايام|شهر|اشهر|شهور|اسبوع|اسابيع)",
    re.IGNORECASE,
)
_DELAI_UNITS = {
    "jour": "JOURS", "jours": "JOURS", "j": "JOURS", "يوم": "JOURS", "ايام": "JOURS",
    "mois": "MOIS", "شهر": "MOIS", "اشهر": "MOIS", "شهور": "MOIS",
    "semaine": "SEMAINES", "semaines": "SEMAINES", "اسبوع": "SEMAINES", "اسابيع": "SEMAINES",
}
# "10 jours à compter de la publication" is the contestation window, not the execution delay
_RECOURS_CONTEXT = re.compile(r"recours|contestation|reclamation|a compter de|طعن|تظلم", re.IGNORECASE)
_EXECUTION_CONTEXT = re.compile(r"delai|duree|execution|realisation|مده|اجل|انجاز", re.IGNORECASE)

def normalize_delai(text) -> str:
    """Execution delay as 'NN JOURS' / 'NN MOIS' / 'NN SEMAINES', or '/'."""
    if text is None:
        return "/"
    t = normalize_text(text)
    best = None
    for m in _DELAI_RE.finditer(t):
        before = t[max(0, m.start() - 60):m.start()]
        if _RECOURS_CONTEXT.search(t[max(0, m.start() - 25):m.start()]) or _RECOURS_CONTEXT.search(t[m.end():m.end() + 40]):
            continue
        value = f"{m.group(1)} {_DELAI_UNITS[m.group(2).lower()]}"
        if _EXECUTION_CONTEXT.search(before):
            return value
        if best is None:
            best = value
    return best or "/"

def parse_delai(nbr_jours: int, description: str) -> str:
    try:
        if nbr_jours and int(nbr_jours) > 0:
            return f"{nbr_jours} JOURS"
    except (TypeError, ValueError):
        pass
    return normalize_delai(description)

ALGERIA_TZ = timezone(timedelta(hours=1))

def to_local_date(value) -> str:
    """
    YYYY-MM-DD in Algerian time. API timestamps are UTC ('2026-10-21T23:00:00.000Z' is the 22nd
    in Algiers), so slicing the first 10 characters can land on the previous day.
    """
    if not value:
        return ""
    s = str(value).strip()
    if re.match(r"^\d{4}-\d{2}-\d{2}T", s) and re.search(r"(?:Z|[+-]\d{2}:?\d{2})$", s):
        try:
            return datetime.fromisoformat(s.replace("Z", "+00:00")).astimezone(ALGERIA_TZ).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return normalize_date_key(s)

def clean_date_parution_ao(value, date_attribution: str) -> str:
    """
    Original AO publication date for an award, as DD/MM/YYYY, or '/'.
    It must be a real date, not after the award itself and not more than two years before it.
    """
    key = normalize_date_key(value)
    if not key:
        return "/"
    try:
        d = datetime.strptime(key, "%Y-%m-%d")
    except ValueError:
        log.warning("Discarding invalid date_parution_ao '%s'", value)
        return "/"
    ref_key = normalize_date_key(date_attribution) or datetime.now().strftime("%Y-%m-%d")
    ref = datetime.strptime(ref_key, "%Y-%m-%d")
    if d > ref or d < ref - timedelta(days=730):
        log.warning("Discarding implausible date_parution_ao '%s' (award date %s)", value, ref_key)
        return "/"
    return d.strftime("%d/%m/%Y")

def parse_date(date_val):
    if not date_val:
        return "/"
    if isinstance(date_val, datetime):
        return datetime(date_val.year, date_val.month, date_val.day)
    try:
        s = str(date_val)[:10]
        dt = datetime.strptime(s, "%Y-%m-%d")
        return datetime(dt.year, dt.month, dt.day)
    except Exception:
        return "/"

# ── AI Scan Vision (NVIDIA NIM / DeepSeek OCR) ────────────────────────────────

NVIDIA_MODEL = "z-ai/glm-5.3-flash"
# (model, read timeout in seconds). moonshotai/kimi-k3 was retired by NVIDIA (HTTP 410, Oct 2026).
NVIDIA_FALLBACK_MODELS = [
    ("z-ai/glm-5.3-flash", 60),
]
# After this many consecutive NVIDIA failures in one run, go straight to Gemini for the rest of it
NVIDIA_MAX_CONSECUTIVE_FAILURES = 3
_nvidia_state = {"failures": 0, "disabled": False}


PROMPT_SCAN_ANALYSIS = """You are an expert document analysis and OCR system specialized in Algerian public procurement (Marchés Publics / الصفقات العمومية) for TAPIDOR.
The document may be in French, Arabic, or bilingual.

TAPIDOR ONLY works on two product families:
1. GAZON — sports or landscape turf: GAZON / PELOUSE (synthétique ou naturel), ENGAZONNEMENT, REVÊTEMENT EN GAZON, POSE DE GAZON, عشب اصطناعي / عشب طبيعي / تعشيب / تكسية بالعشب الاصطناعي.
2. TAPIS — carpets and moquette for MOSQUES or BUILDINGS (offices, schools, hotels, conference rooms): TAPIS, MOQUETTE, زرابي / زربية / سجاد / موكيت.

NOT IN SCOPE (scope must be "NONE"):
- Projects that only mention 'stade', 'terrain de sport', 'ملعب', 'aire de jeux', 'complexe sportif', 'peinture', 'éclairage', 'clôture', 'gradins', 'vestiaires' without explicit gazon/pelouse/turf.
- Civil masonry, school rehabilitation, roadworks, concrete/asphalt/MATICO multi-sports grounds WITHOUT turf.
- Tondeuses à gazon / lawn mowers, matériel agricole, tonte d'espaces verts.
- "Tapis" that is NOT a carpet: tapis roulant (treadmill/conveyor), tapis d'enrobé / tapis en béton bitumineux (road asphalt layer), tatami, tapis de judo / gymnastique / sol sportif, tapis de souris.

STRICT RULE: set "scope" to "GAZON" or "TAPIS" ONLY if the document EXPLICITLY names the product, and quote it verbatim in "scope_evidence". Otherwise "scope": "NONE" and "scope_evidence": "NONE".

LOTS (very important for award notices / avis d'attribution):
If the notice lists several lots (Lot 01, Lot 02, الحصة 01 ...), return EACH lot separately in "lots" with ITS OWN amount, winner and delay. NEVER add lot amounts together and NEVER merge several lots into one budget.
Set "in_scope": true only for lots whose object explicitly concerns gazon or tapis/moquette.
If there is a single lot or no lot structure, return one entry in "lots".

DATE DE PARUTION REQUIREMENT (FOR AVIS D'ATTRIBUTION):
Extract "date_parution_ao" ONLY if the text explicitly states the date when the original Appel d'Offres was published in national newspapers (e.g. 'paru dans les journaux le DD/MM/2026' or 'الصادرة بتاريخ 2026/MM/DD في جريدة ...').
Format: DD/MM/YYYY.
CRITICAL WARNING - NEVER INVENT A DATE:
If no newspaper publication date is explicitly mentioned in the text (such as for consultations / استشارة, or direct awards), you MUST set "date_parution_ao": "/".
NEVER use the newspaper print date or footer date at the very bottom (such as 'An-Nasr 21-9-2026', ANEP footer, or 'El-Moudjahid').
NEVER use dates from decrees (e.g. 2015, 2023), dates of consultation, or date of bid opening.

AMOUNTS: copy exactly ONE amount per field, digits as printed, followed by DA (e.g. '37 756 320,00 DA'). Prefer the final corrected TTC amount. Never put a date or a lot number in an amount field.

Extract the following in strict JSON:
{
  "scope": "GAZON" or "TAPIS" or "NONE",
  "scope_evidence": "Exact verbatim quote naming gazon/pelouse/عشب or tapis/moquette/زرابي, or 'NONE'.",
  "objet_marche": "Verbatim title or project object from the document in French or Arabic",
  "action": "Main action verb in French (e.g. REALISATION, AMENAGEMENT, REVETEMENT, REHABILITATION, FOURNITURE ET POSE) or '/'",
  "type_projet": "Facility type in French (e.g. STADE, STADE COMMUNAL, STADE DE PROXIMITE, AIRE DE JEUX, TERRAIN DE SPORT, MOSQUEE) or '/'",
  "budget": "Single winning amount TTC (only when there is ONE lot) or '/'",
  "delai": "Execution delay 'XX JOURS' or 'XX MOIS' (only when there is ONE lot) or '/'",
  "commune": "Commune (municipality) name in UPPERCASE Latin letters or '/'",
  "wilaya": "Wilaya name in UPPERCASE or '/'",
  "annonceur": "Contracting authority name in French or '/'",
  "entreprise_attributaire": "Winning contractor in Latin letters (only when there is ONE lot) or '/'",
  "date_parution_ao": "Original Appel d'Offres publication date from INSIDE the paragraph text, DD/MM/YYYY, or '/'",
  "lots": [
    {
      "numero": "01",
      "objet": "Verbatim object of this lot",
      "in_scope": true or false,
      "commune": "Commune of this lot in UPPERCASE or '/'",
      "budget": "Amount awarded for THIS lot only, e.g. '6 959 358,00 DA', or '/'",
      "entreprise_attributaire": "Winner of THIS lot or '/'",
      "delai": "Delay of THIS lot, e.g. '45 JOURS', or '/'"
    }
  ]
}
Return ONLY valid JSON."""

def prepare_image_for_ocr(img_bytes: bytes, max_dim: int = 1600) -> tuple[bytes, str]:
    """Ensure image is optimized to prevent timeout on vision endpoints."""
    if not img_bytes:
        return b"", "image/jpeg"
    if img_bytes[:4] == b"%PDF":
        return img_bytes, "application/pdf"
    mime = "image/png" if img_bytes[:8] == b"\x89PNG\r\n\x1a\n" else "image/jpeg"
    try:
        from PIL import Image
        import io
        im = Image.open(io.BytesIO(img_bytes))
        if im.mode in ("RGBA", "P"):
            im = im.convert("RGB")
        if max(im.size) > max_dim or len(img_bytes) > 350 * 1024:
            im.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
            buf = io.BytesIO()
            im.save(buf, format="JPEG", quality=85, optimize=True)
            return buf.getvalue(), "image/jpeg"
    except Exception as e:
        log.debug("Image optimization bypassed: %s", e)
    return img_bytes, mime

def extract_json_from_text(text: str) -> dict:
    if not text:
        return {}
    try:
        return json.loads(text.strip())
    except Exception:
        pass
    m = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", text)
    if m:
        try:
            return json.loads(m.group(1).strip())
        except Exception:
            pass
    m = re.search(r"(\{[\s\S]*\})", text)
    if m:
        try:
            return json.loads(m.group(1).strip())
        except Exception:
            pass
    return {}

def analyze_scan_with_nvidia(img_bytes: bytes, api_key: str | None = None) -> dict:
    """
    Extract complementary details from attached ad scans using NVIDIA NIM.
    Primary model: z-ai/glm-5.3-flash. Returns the raw model JSON, or {} on failure.
    """
    key = api_key or NVIDIA_API_KEY
    if not key or not img_bytes:
        return {}

    opt_bytes, mime = prepare_image_for_ocr(img_bytes)
    if mime == "application/pdf":
        # NIM chat models only accept images; PDFs go to Gemini
        return {}
    b64_img = base64.b64encode(opt_bytes).decode("utf-8")

    url = "https://integrate.api.nvidia.com/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json"
    }

    for model, timeout in NVIDIA_FALLBACK_MODELS:
        payload = {
            "model": model,
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": PROMPT_SCAN_ANALYSIS},
                        {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64_img}"}}
                    ]
                }
            ],
            "max_tokens": 2048,
            "temperature": 0.1
        }
        try:
            r = requests.post(url, headers=headers, json=payload, timeout=timeout)
            if r.status_code == 200:
                res = r.json()
                msg = res["choices"][0]["message"]
                raw_text = msg.get("content") or msg.get("reasoning_content") or ""
                data = extract_json_from_text(raw_text)
                if data and isinstance(data, dict):
                    log.info("NVIDIA NIM (%s) extracted -> scope: %s | Commune: %s | Budget: %s | Lots: %d | Parution AO: %s",
                             model, data.get("scope", data.get("is_gazon")), data.get("commune"), data.get("budget"),
                             len(data.get("lots") or []), data.get("date_parution_ao"))
                    time.sleep(0.5)
                    return data
                log.warning("NVIDIA NIM (%s) answered without parseable JSON: %s", model, raw_text[:120])
            elif r.status_code in (429, 503):
                log.warning("NVIDIA NIM (%s) status %d -- falling back...", model, r.status_code)
            else:
                log.warning("NVIDIA NIM (%s) returned %d: %s", model, r.status_code, r.text[:120])
        except Exception as e:
            log.warning("NVIDIA NIM (%s) request error (%s) -- falling back...", model, e)

    return {}

# gemini-2.0-flash and gemini-1.5-flash return 404 since Oct 2026
GEMINI_MODELS = [
    "gemini-2.5-flash",
    "gemini-flash-latest",
]
GEMINI_TIMEOUT = 60
GEMINI_RETRY_DELAYS = (5, 15)  # seconds to wait before retrying a model after 429/503/timeout

def analyze_scan_with_gemini(img_bytes: bytes, api_key: str | None = None) -> dict:
    key = api_key or GEMINI_API_KEY
    if not key or not img_bytes:
        return {}

    opt_bytes, mime = prepare_image_for_ocr(img_bytes)
    b64_img = base64.b64encode(opt_bytes).decode("utf-8")

    payload = {
        "contents": [
            {
                "parts": [
                    {"text": PROMPT_SCAN_ANALYSIS},
                    {"inline_data": {"mime_type": mime, "data": b64_img}}
                ]
            }
        ],
        "generationConfig": {"response_mime_type": "application/json", "temperature": 0.1}
    }

    for model in GEMINI_MODELS:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={key}"
        for attempt in range(len(GEMINI_RETRY_DELAYS) + 1):
            retryable = False
            try:
                r = requests.post(url, json=payload, headers={"Content-Type": "application/json"}, timeout=GEMINI_TIMEOUT)
                if r.status_code == 200:
                    content = r.json()["candidates"][0]["content"]["parts"][0]["text"]
                    data = extract_json_from_text(content)
                    if data and isinstance(data, dict):
                        log.info("Gemini Vision (%s) extracted -> scope: %s | Commune: %s | Budget: %s | Lots: %d | Parution AO: %s",
                                 model, data.get("scope", data.get("is_gazon")), data.get("commune"), data.get("budget"),
                                 len(data.get("lots") or []), data.get("date_parution_ao"))
                        time.sleep(0.5)
                        return data
                    log.warning("Gemini (%s) answered without parseable JSON: %s", model, content[:120])
                elif r.status_code in (429, 500, 503):
                    retryable = True
                    log.warning("Gemini (%s) status %d (attempt %d)", model, r.status_code, attempt + 1)
                else:
                    log.warning("Gemini (%s) returned %d: %s", model, r.status_code, r.text[:120])
            except requests.exceptions.Timeout as e:
                retryable = True
                log.warning("Gemini (%s) timed out (attempt %d): %s", model, attempt + 1, e)
            except Exception as e:
                log.warning("Gemini (%s) request error (%s) -- falling back...", model, e)
            if not retryable or attempt >= len(GEMINI_RETRY_DELAYS):
                break
            time.sleep(GEMINI_RETRY_DELAYS[attempt])
    return {}


def analyze_scan(img_bytes: bytes) -> dict:
    """
    Read one scan: NVIDIA NIM (GLM 5.3 Flash) first, then Gemini 2.5 Flash / flash-latest.
    Returns normalize_ai_data() output, or {} when every engine failed (scan could not be read).
    """
    if NVIDIA_API_KEY and not _nvidia_state["disabled"]:
        try:
            data = analyze_scan_with_nvidia(img_bytes)
        except Exception as e:
            log.warning("NVIDIA NIM analysis error: %s", e)
            data = {}
        if data:
            _nvidia_state["failures"] = 0
            return normalize_ai_data(data)
        if img_bytes[:4] != b"%PDF":
            _nvidia_state["failures"] += 1
            if _nvidia_state["failures"] >= NVIDIA_MAX_CONSECUTIVE_FAILURES:
                _nvidia_state["disabled"] = True
                log.warning("NVIDIA NIM failed %d times in a row -- using Gemini only for the rest of this run",
                            _nvidia_state["failures"])

    if GEMINI_API_KEY:
        try:
            return normalize_ai_data(analyze_scan_with_gemini(img_bytes))
        except Exception as e:
            log.warning("Gemini fallback analysis error: %s", e)

    return {}

# ── AI output normalization ──────────────────────────────────────────────────

AI_TEXT_FIELDS = (
    "objet_marche", "scope_evidence", "action", "type_projet", "budget", "delai", "commune",
    "wilaya", "annonceur", "entreprise_attributaire", "date_parution_ao",
)
AI_LOT_FIELDS = ("numero", "objet", "commune", "budget", "entreprise_attributaire", "delai")

def _ai_str(v) -> str:
    """Models return null, numbers, lists or booleans in text fields; turn all of them into text or '/'."""
    if v is None or isinstance(v, bool):
        return "/"
    if isinstance(v, (int, float)):
        return str(v)
    if isinstance(v, list):
        return ", ".join(s for s in (_ai_str(x) for x in v) if s != "/") or "/"
    if isinstance(v, dict):
        return "/"
    s = str(v).strip()
    return s if s and s.upper() not in ("NULL", "NONE", "N/A") else "/"

def _ai_bool(v) -> bool:
    return v is True or str(v).strip().lower() in ("true", "yes", "oui", "1")

def normalize_ai_data(raw) -> dict:
    """Coerce a model answer into the expected shape. {} stays {} (meaning: OCR failed)."""
    if not isinstance(raw, dict) or not raw:
        return {}
    data = {k: _ai_str(raw.get(k)) for k in AI_TEXT_FIELDS}
    if data["scope_evidence"] == "/":
        data["scope_evidence"] = _ai_str(raw.get("gazon_mention"))

    scope = str(raw.get("scope") or "").strip().upper()
    if scope not in ("GAZON", "TAPIS", "NONE"):
        scope = "GAZON" if _ai_bool(raw.get("is_gazon")) else "NONE"
    data["scope"] = scope

    lots = []
    for lot in raw.get("lots") or []:
        if not isinstance(lot, dict):
            continue
        clean = {k: _ai_str(lot.get(k)) for k in AI_LOT_FIELDS}
        clean["in_scope"] = _ai_bool(lot.get("in_scope"))
        lots.append(clean)
    data["lots"] = lots
    return data

def merge_ai_data(primary: dict, extra: dict) -> dict:
    """Combine two pages of the same notice: fill gaps from `extra`, keep any in-scope verdict."""
    if not primary:
        return extra or {}
    if not extra:
        return primary
    merged = dict(primary)
    for k in AI_TEXT_FIELDS:
        if merged.get(k, "/") == "/" and extra.get(k, "/") != "/":
            merged[k] = extra[k]
    if merged.get("scope") not in ("GAZON", "TAPIS") and extra.get("scope") in ("GAZON", "TAPIS"):
        merged["scope"] = extra["scope"]
        merged["scope_evidence"] = extra.get("scope_evidence", "/")
    if not merged.get("lots") and extra.get("lots"):
        merged["lots"] = extra["lots"]
    return merged

def ai_confirmed_scope(ai_data: dict) -> str | None:
    """The scope the AI claims, but only if its quoted evidence really names gazon or tapis."""
    if not ai_data or ai_data.get("scope") not in ("GAZON", "TAPIS"):
        return None
    return detect_scope(ai_data.get("scope_evidence", ""))

# ── Google Drive Manager ─────────────────────────────────────────────────────

class GoogleDriveManager:
    """Manages date-folder organization and scan uploads in Google Drive."""

    def __init__(self, folder_id: str | None = None):
        self.folder_id = folder_id or GOOGLE_DRIVE_FOLDER_ID
        self.service = None
        self.quota_exceeded = False
        self._date_folders = {}

        if not self.folder_id:
            return

        creds = None
        sa_json = os.getenv("GCP_SERVICE_ACCOUNT_KEY")
        sa_file = SCRIPT_DIR / "state" / "service_account.json"
        try:
            if sa_json:
                try:
                    if not sa_json.strip().startswith("{"):
                        sa_json = base64.b64decode(sa_json).decode("utf-8")
                    creds_dict = json.loads(sa_json)
                    if "private_key" in creds_dict and "\\n" in creds_dict["private_key"]:
                        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
                    creds = service_account.Credentials.from_service_account_info(
                        creds_dict,
                        scopes=["https://www.googleapis.com/auth/drive"]
                    )
                except Exception as e:
                    log.warning("Drive credentials error: %s", e)
            elif sa_file.exists():
                creds = service_account.Credentials.from_service_account_file(
                    str(sa_file),
                    scopes=["https://www.googleapis.com/auth/drive"]
                )
            if creds:
                self.service = build("drive", "v3", credentials=creds)
                log.info("Google Drive service initialized (Target Folder: %s)", self.folder_id)
        except Exception as e:
            log.warning("Could not initialize Google Drive service: %s", e)

    def get_or_create_date_folder(self, date_str: str) -> str | None:
        if not self.service or not self.folder_id:
            return None
        if date_str in self._date_folders:
            return self._date_folders[date_str]

        try:
            q = f"'{self.folder_id}' in parents and name = '{date_str}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
            res = self.service.files().list(q=q, fields="files(id, name)").execute()
            files = res.get("files", [])
            if files:
                self._date_folders[date_str] = files[0]["id"]
                return files[0]["id"]

            meta = {
                "name": date_str,
                "mimeType": "application/vnd.google-apps.folder",
                "parents": [self.folder_id]
            }
            folder = self.service.files().create(body=meta, fields="id, name").execute()
            fid = folder.get("id")
            self._date_folders[date_str] = fid
            log.info("Created Google Drive date folder '%s' (ID: %s)", date_str, fid)
            return fid
        except Exception as e:
            log.warning("Google Drive: unable to get/create date folder %s: %s", date_str, e)
            return None

    def upload_scan(self, file_path: Path, date_str: str) -> dict[str, str] | None:
        if not self.service or self.quota_exceeded or not file_path.exists():
            return None

        date_folder_id = self.get_or_create_date_folder(date_str)
        if not date_folder_id:
            return None

        try:
            q = f"'{date_folder_id}' in parents and name = '{file_path.name}' and trashed = false"
            res = self.service.files().list(q=q, fields="files(id, name, webViewLink)").execute()
            existing = res.get("files", [])
            if existing:
                return {"id": existing[0]["id"], "url": existing[0].get("webViewLink", "")}

            mime = "image/jpeg" if file_path.suffix.lower() in [".jpg", ".jpeg"] else "application/octet-stream"
            media = MediaFileUpload(str(file_path), mimetype=mime, resumable=True)
            file_meta = {
                "name": file_path.name,
                "parents": [date_folder_id]
            }
            uploaded = self.service.files().create(body=file_meta, media_body=media, fields="id, name, webViewLink").execute()
            log.info("Uploaded scan %s to Google Drive (%s)", file_path.name, uploaded["id"])
            return {"id": uploaded["id"], "url": uploaded.get("webViewLink", "")}
        except HttpError as e:
            if "storageQuotaExceeded" in str(e):
                self.quota_exceeded = True
                log.warning(
                    "Drive upload notice: Service Account personal storage quota is 0 MB (Google limitation for personal @gmail folders). "
                    "Scans are safely stored in GitHub and local data/scans. To enable direct Drive upload, use a Google Workspace Shared Drive."
                )
            else:
                log.warning("Drive upload failed for %s: %s", file_path.name, e)
            return None
        except Exception as e:
            log.warning("Drive upload error for %s: %s", file_path.name, e)
            return None




# ── Google Sheets Sync ───────────────────────────────────────────────────────

def normalize_date_key(d: str) -> str:
    """Normalize date string to YYYY-MM-DD for accurate cross-format comparison."""
    if not d or d == "/":
        return ""
    m = re.search(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', str(d))
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    m = re.search(r'(\d{1,2})[-/](\d{1,2})[-/](\d{4})', str(d))
    if m:
        return f"{int(m.group(3)):04d}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"
    return ""

def format_date_dmy(d_str: str) -> str:
    """Format date string as DD/MM/YYYY for French/Algerian locale Google Sheets."""
    if not d_str or d_str == "/":
        return "/"
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", str(d_str).strip())
    if m:
        return f"{m.group(3)}/{m.group(2)}/{m.group(1)}"
    return str(d_str).strip()

def sheet_text(value) -> str:
    """
    Force Google Sheets to keep a value as literal text. The spreadsheet locale is en_US, so a
    USER_ENTERED '02/11/2026' would silently become February 11 (shown as serial 46064).
    """
    s = str(value).strip() if value is not None else ""
    if not s or s == "/":
        return "/"
    return "'" + s

GITHUB_REPO_URL = "https://github.com/bronsonclaudia3-png/algeriemarches-scraper"

# Notice IDs are embedded in every SCANS cell: in the GitHub scan folder path, or as #ad-<id>
# on the AlgerieMarches link written when a notice has no scan.
AD_ID_PATTERNS = (
    re.compile(r'/(\d{5,6})(?:/|"|\.|$)'),
    re.compile(r'#ad-(\d{5,6})\b'),
)

def extract_ad_ids(cell) -> set[str]:
    found = set()
    for pattern in AD_ID_PATTERNS:
        found.update(pattern.findall(str(cell)))
    return found

def _norm_key(s) -> str:
    return re.sub(r"[^A-Z0-9]", "", strip_accents(str(s or "")).upper())

def _budget_key(value) -> str:
    v = parse_budget(value)
    return f"{float(v):.2f}" if v != "/" else ""

def build_scan_link(item: dict) -> str:
    ann_id = str(item.get("id", "") or item.get("id_annonce", "")).strip()
    dt_parution = to_local_date(item.get("date_parution")) or "/"
    scan_names = [Path(s).name for s in item.get("scans") or []]
    if scan_names and ann_id and dt_parution != "/":
        if len(scan_names) == 1:
            return f'=HYPERLINK("{GITHUB_REPO_URL}/blob/main/data/scans/{dt_parution}/{ann_id}/{scan_names[0]}", "📄 Voir Scan")'
        return f'=HYPERLINK("{GITHUB_REPO_URL}/tree/main/data/scans/{dt_parution}/{ann_id}", "📁 Voir {len(scan_names)} Scans")'
    if ann_id and item.get("slug"):
        return f'=HYPERLINK("{BASE_URL}/annonces/{item["slug"]}#ad-{ann_id}", "🔗 Annonce")'
    return "/"

def ao_dedup_keys(action, ptype, norm_w, norm_c, parution_key, echeance_key) -> list[tuple]:
    """Secondary AO keys. Never use a date-less location key (it dropped notices 85018 and 93574)."""
    if not norm_w:
        return []
    keys = []
    if parution_key:
        keys.append(("AO_PARUTION", action, ptype, norm_w, norm_c, parution_key))
    if echeance_key:
        keys.append(("AO_ECHEANCE", norm_w, norm_c, echeance_key))
    if not parution_key and not echeance_key:
        keys.append(("AO_NODATE", action, ptype, norm_w, norm_c))
    return keys

def aa_dedup_key(norm_w, norm_c, budget_key, attr_norm, date_attr_key) -> tuple | None:
    """Secondary AA key for one lot row: the exact amount in that commune, else winner + award date."""
    if not norm_w:
        return None
    if budget_key:
        return ("AA_BUDGET", norm_w, norm_c, budget_key)
    if attr_norm and date_attr_key:
        return ("AA_NOBUDGET", norm_w, norm_c, attr_norm, date_attr_key)
    return None

def read_sheet_state(ws, notice_type: str) -> dict | None:
    """Locate the data columns and index what the worksheet already contains."""
    try:
        all_values = ws.get_all_values()
        all_formulas = ws.get_all_values(value_render_option="FORMULA")
    except Exception as e:
        log.error("Failed to read values from Google Sheet: %s", e)
        return None

    header_row = all_values[2] if len(all_values) > 2 else []
    col_offset = next((i for i, cell in enumerate(header_row) if str(cell).strip().upper() == "N°"), None)
    if col_offset is None:
        log.error("Worksheet '%s': no 'N°' header in row 3 -- refusing to write misaligned rows", ws.title)
        return None
    scans_idx = col_offset + (11 if notice_type == "appels-doffres" else 15)
    scans_header = str(header_row[scans_idx]).strip().upper() if len(header_row) > scans_idx else ""
    if scans_header != "SCANS":
        log.error("Worksheet '%s': expected 'SCANS' header at column %d, found '%s' -- layout changed, not writing",
                  ws.title, scans_idx + 1, scans_header)
        return None

    def cell(row, i):
        return str(row[col_offset + i]).strip() if len(row) > col_offset + i else ""

    last_num = 0
    ids, items = set(), set()
    for idx, (row, frow) in enumerate(zip(all_values, all_formulas)):
        if idx < 3:  # Skip title & header rows
            continue
        for c in list(row) + list(frow):
            ids.update(extract_ad_ids(c))
        try:
            last_num = max(last_num, int(cell(row, 0)))
        except ValueError:
            pass

        norm_w = _norm_key(cell(row, 8))
        norm_c = _norm_key(cell(row, 9))
        if notice_type == "appels-doffres":
            parution = normalize_date_key(cell(row, 5)) or normalize_date_key(cell(row, 1))
            items.update(ao_dedup_keys(cell(row, 2).upper(), cell(row, 4).upper(), norm_w, norm_c,
                                       parution, normalize_date_key(cell(row, 6))))
        else:
            key = aa_dedup_key(norm_w, norm_c, _budget_key(cell(row, 11)), _norm_key(cell(row, 12)),
                               normalize_date_key(cell(row, 10)))
            if key:
                items.add(key)

    return {"col_offset": col_offset, "last_num": last_num, "ids": ids, "items": items, "n_rows": len(all_values)}

def build_ao_row(item: dict, today_str: str) -> tuple[list, list[tuple]]:
    """One APPEL D'OFFRE row (without N°) and its dedup keys."""
    titre = item.get("titre", "")
    action = item.get("action") or classify_action(titre)
    ptype = item.get("ptype") or classify_type_projet(titre)
    wilaya = (item.get("wilaya") or "").strip().upper()
    commune = item.get("commune") if item.get("commune") not in (None, "", "/") else parse_commune(titre, item.get("annonceur", ""))
    dt_parution = to_local_date(item.get("date_parution")) or "/"
    dt_echeance = to_local_date(item.get("date_echeance")) or "/"

    row = [
        sheet_text(format_date_dmy(dt_parution if dt_parution != "/" else today_str)),  # DATE
        action,                                     # TITRE D'APPEL D'OFFRE
        1,                                          # Nombre de projet
        ptype,                                      # TYPE DE PROJET
        sheet_text(format_date_dmy(dt_parution)),   # DATE DE PARUTION
        sheet_text(format_date_dmy(dt_echeance)),   # DATE D'ECHEANCE
        clean_annonceur(item.get("annonceur", "")),  # ANNONCEUR
        wilaya or "/",                              # WILAYA
        commune,                                    # COMMUNE
        "TRAVAUX PUBLICS",                          # CATEGORIE
        build_scan_link(item),                      # SCANS
    ]
    keys = ao_dedup_keys(action.upper(), ptype.upper(), _norm_key(wilaya), _norm_key(commune),
                         normalize_date_key(dt_parution), normalize_date_key(dt_echeance))
    return row, keys

GENERIC_PTYPES = {"INFRASTRUCTURE SPORTIVE", "TERRAIN", "COUR"}

def build_aa_rows(item: dict, today_str: str) -> list[tuple[list, tuple | None]]:
    """One AVIS D'ATTRIBUTION row (without N°) per in-scope lot, each with its own amount and winner."""
    titre = item.get("titre", "")
    action = item.get("action") or classify_action(titre)
    ptype = item.get("ptype") or classify_type_projet(titre)
    wilaya = (item.get("wilaya") or "").strip().upper()
    commune = item.get("commune") if item.get("commune") not in (None, "", "/") else parse_commune(titre, item.get("annonceur", ""))
    dt_parution = to_local_date(item.get("date_parution")) or "/"
    dt_echeance = to_local_date(item.get("date_echeance")) or "/"
    date_col_val = dt_parution if dt_parution != "/" else today_str
    dt_parution_ao = clean_date_parution_ao(item.get("date_parution_ao"), date_col_val)
    ann_val = clean_annonceur(item.get("annonceur", ""))
    scan_link = build_scan_link(item)
    item_delai = item.get("delai") or parse_delai(item.get("nbr_jours", 0), item.get("description", ""))

    lots = item.get("lots") or [{}]
    single = len(lots) == 1
    rows = []
    for lot in lots:
        objet = lot.get("objet", "/")
        lot_action = classify_action(objet) if objet != "/" else "TRAVAUX"
        lot_ptype = classify_type_projet(objet) if objet != "/" else "INFRASTRUCTURE SPORTIVE"
        lot_commune = lot.get("commune", "/")
        lot_commune = strip_accents(lot_commune).upper().strip() if lot_commune != "/" else commune
        budget = parse_budget(lot.get("budget"))
        if budget == "/" and single:
            # The notice-level amount only belongs to the row when there is a single lot;
            # with several lots it is a total and must never be copied onto one lot.
            budget = parse_budget(item.get("montant"))
        attr_par = lot.get("entreprise_attributaire", "/")
        if attr_par == "/":
            attr_par = (item.get("entreprise_concernee") or "").strip() or "/"
        delai = normalize_delai(lot.get("delai")) if lot.get("delai", "/") != "/" else item_delai

        row = [
            sheet_text(date_col_val),                                   # DATE
            lot_action if lot_action != "TRAVAUX" else action,          # TITRE D'AVIS D'ATTRIBUTION
            1,                                                          # Nombre de projet
            lot_ptype if lot_ptype not in GENERIC_PTYPES else ptype,    # TYPE DE PROJET
            sheet_text(dt_parution_ao),                                 # DATE DE PARUTION (AO, from the paragraph)
            sheet_text(dt_echeance),                                    # DATE D'ECHEANCE
            ann_val,                                                    # ANNONCEUR
            wilaya or "/",                                              # WILAYA
            lot_commune,                                                # COMMUNE
            sheet_text(dt_parution),                                    # DATE D'ATTRIBUTION
            budget,                                                     # BUDGET
            attr_par,                                                   # ATTRIBUTION PAR
            delai or "/",                                               # DELAI
            "/",                                                        # WILAYA 2
            scan_link,                                                  # SCANS
        ]
        key = aa_dedup_key(_norm_key(wilaya), _norm_key(lot_commune), _budget_key(budget),
                           _norm_key(attr_par if attr_par != "/" else ""), normalize_date_key(dt_parution))
        rows.append((row, key))
    return rows

def open_worksheet(notice_type: str, sheet_id: str | None = None):
    """Return (worksheet, status) where status is 'ok', 'no_creds' or 'error'."""
    sheet_id = sheet_id or os.getenv("GOOGLE_SHEET_ID", "1y5-OxNeL_zBhCUNNEVoh1nKZVsyBYrq5EJcgKBEt920")
    if not sheet_id:
        log.info("No GOOGLE_SHEET_ID configured -- skipping Google Sheet sync.")
        return None, "no_creds"

    sa_json_env = os.getenv("GCP_SERVICE_ACCOUNT_KEY")
    sa_file = SCRIPT_DIR / "state" / "service_account.json"

    gc = None
    if sa_json_env:
        try:
            if not sa_json_env.strip().startswith("{"):
                sa_json_env = base64.b64decode(sa_json_env).decode("utf-8")
            creds_dict = json.loads(sa_json_env)
            if "private_key" in creds_dict and "\\n" in creds_dict["private_key"]:
                creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
            gc = gspread.service_account_from_dict(creds_dict)
            log.info("Authenticated with Google via GCP_SERVICE_ACCOUNT_KEY env var")
        except Exception as e:
            log.error("Failed to authenticate with GCP_SERVICE_ACCOUNT_KEY: %s", e)
            return None, "error"
    elif sa_file.exists():
        try:
            gc = gspread.service_account(filename=str(sa_file))
            log.info("Authenticated with Google via state/service_account.json")
        except Exception as e:
            log.error("Failed to authenticate with service_account.json: %s", e)
            return None, "error"

    if not gc:
        log.info("No Google service account credentials found -- skipping Google Sheet sync.")
        return None, "no_creds"

    try:
        sh = gc.open_by_key(sheet_id)
        log.info("Opened Google Sheet: %s", sh.title)
    except Exception as e:
        log.error("Could not access Google Sheet '%s': %s", sheet_id, e)
        log.error(">> Please ensure you clicked 'Share' on the Google Sheet and added:")
        log.error(">> algeriemarches-bot@project-df86d806-9e9d-42be-a7d.iam.gserviceaccount.com as Editor!")
        return None, "error"

    target_title = "APPEL D'OFFRE 2026" if notice_type == "appels-doffres" else "AVIS D'ATTRIBUTIONS 2026"
    for ws in sh.worksheets():
        if ws.title.strip().upper() == target_title.upper():
            log.info("Target Google Sheet worksheet: %s", ws.title)
            return ws, "ok"
    log.error("Worksheet '%s' not found in Google Sheet", target_title)
    return None, "error"

def append_to_gsheet(results: list[dict], notice_type: str, sheet_id: str | None = None, worksheet=None) -> int:
    """
    Write new results to the matching worksheet in Google Sheet.
    - notice_type 'appels-doffres' -> 'APPEL D'OFFRE 2026' (one row per notice, data from column C)
    - notice_type 'avis-attribution' -> 'AVIS D'ATTRIBUTIONS 2026' (one row per in-scope lot, from column A)
    Returns the number of rows written, 0 when Sheets is not configured, -1 on failure.
    """
    ws = worksheet
    if ws is None:
        ws, status = open_worksheet(notice_type, sheet_id)
        if ws is None:
            return 0 if status == "no_creds" else -1

    state = read_sheet_state(ws, notice_type)
    if state is None:
        return -1
    log.info("Google Sheet '%s' has %d rows (last N°: %d, tracked IDs: %d)",
             ws.title, state["n_rows"] - 3, state["last_num"], len(state["ids"]))

    today_str = datetime.now().strftime("%Y-%m-%d")
    last_num = state["last_num"]
    new_rows = []

    for item in results:
        ann_id = str(item.get("id", "") or item.get("id_annonce", "")).strip()
        if ann_id and ann_id in state["ids"]:
            log.info("  Notice ID %s already in Google Sheet '%s' -- skipping duplicate!", ann_id, ws.title)
            continue

        if notice_type == "appels-doffres":
            row, ad_keys = build_ao_row(item, today_str)
            built = [(row, None)]
        else:
            built = build_aa_rows(item, today_str)
            ad_keys = [k for _, k in built if k]

        # Deduplicate per notice: lots of the same award never block each other
        if any(k in state["items"] for k in ad_keys):
            log.info("  Already in Google Sheet: %s (%s, %s)", item.get("titre", "")[:50], item.get("wilaya"), item.get("commune"))
            continue

        for row, _ in built:
            last_num += 1
            new_rows.append([last_num] + row)
        state["items"].update(ad_keys)
        if ann_id:
            state["ids"].add(ann_id)

    if not new_rows:
        log.info("No new rows needed for Google Sheet '%s' (all up to date)", ws.title)
        return 0

    # Write to an explicit range starting at the N° column. values.append() detected the table at
    # column C and then shifted the blank-padded AO rows two columns right (row 259, Oct 2026).
    start_row = state["n_rows"] + 1
    end_row = start_row + len(new_rows) - 1
    try:
        if end_row > ws.row_count:
            ws.add_rows(end_row - ws.row_count)
        ws.update(values=new_rows, range_name=rowcol_to_a1(start_row, state["col_offset"] + 1),
                  value_input_option="USER_ENTERED")
        log.info("Successfully wrote %d new rows to Google Sheet '%s' (rows %d-%d)", len(new_rows), ws.title, start_row, end_row)
    except Exception as e:
        log.error("Failed to write rows to Google Sheet: %s", e)
        return -1
    return len(new_rows)

# ── Daily Scans Summary Generator ────────────────────────────────────────────

def generate_daily_scans_summary(all_results: list[dict]) -> list[Path]:
    """Generates clean visual Markdown galleries in data/scans/YYYY-MM-DD/README.md for GitHub."""
    by_date: dict[str, list[dict]] = {}
    today_str = datetime.now().strftime("%Y-%m-%d")
    for r in all_results:
        d = to_local_date(r.get("date_parution")) or today_str
        by_date.setdefault(d, []).append(r)

    generated = []
    for d_str, matches in by_date.items():
        day_dir = SCANS_DIR / d_str
        day_dir.mkdir(parents=True, exist_ok=True)
        md_file = day_dir / "README.md"
        lines = [
            f"# 📰 Newspaper Scans Gallery — {d_str}\n",
            "> Scans downloaded & OCR-analyzed (NVIDIA NIM / Gemini) for TAPIDOR\n",
            "| N° | Ad ID | Type | Action & Facility | Wilaya | Commune | Budget | Délai | Scans |",
            "| :---: | :---: | :---: | :--- | :---: | :---: | :---: | :---: | :---: |"
        ]
        for idx, r in enumerate(matches, start=1):
            scans = r.get("scans", [])
            scan_links = []
            for s_path in scans:
                fname = Path(s_path).name
                rel_link = f"{r.get('id')}/{fname}"
                scan_links.append(f"[{fname}]({rel_link})")
            scan_col = "<br>".join(scan_links) if scan_links else "Aucun scan"
            ad_type = "Appel d'Offres" if r.get("type_annonce") == "appels-doffres" else "Avis d'Attribution"
            lines.append(
                f"| {idx} | [{r.get('id')}]({r.get('url')}) | {ad_type} | **{r.get('action')}**<br>{r.get('ptype')} | {r.get('wilaya')} | {r.get('commune')} | {r.get('montant') or '/'} | {r.get('delai') or '/'} | {scan_col} |"
            )

        md_file.write_text("\n".join(lines), encoding="utf-8")
        log.info("Generated daily scans gallery: %s (%d ads)", md_file.relative_to(SCRIPT_DIR), len(matches))
        generated.append(md_file)

    # Also update master data/scans/README.md
    try:
        date_dirs = [d for d in SCANS_DIR.iterdir() if d.is_dir() and re.match(r"^\d{4}-\d{2}-\d{2}$", d.name)]
        date_dirs.sort(key=lambda d: d.name, reverse=True)
        root_lines = [
            "# 🗂️ AlgerieMarches Daily Scans Archive\n",
            "> Central archive of newspaper scans downloaded and OCR-analyzed (NVIDIA NIM / Gemini) for TAPIDOR.\n",
            "| Date | Matched Ads | Gallery Link |",
            "| :---: | :---: | :---: |"
        ]
        for d in date_dirs:
            ad_folders = [f for f in d.iterdir() if f.is_dir()]
            count = len(ad_folders)
            root_lines.append(f"| **{d.name}** | {count} ad{'s' if count != 1 else ''} | [📂 Browse Scans]({d.name}/) |")
        (SCANS_DIR / "README.md").write_text("\n".join(root_lines), encoding="utf-8")
    except Exception as e:
        log.warning("Could not update root scans README: %s", e)

    return generated

# ── Verdict cache ────────────────────────────────────────────────────────────

def load_verdicts(path: Path = VERDICTS_FILE) -> dict:
    """{notice_id: {"verdict": "match"|"reject", "v": VERDICTS_VERSION, "checked": date, ...}}"""
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except FileNotFoundError:
        return {}
    except Exception as e:
        log.warning("Could not read %s (%s) -- starting with an empty verdict cache", path.name, e)
        return {}

def save_verdicts(verdicts: dict, path: Path = VERDICTS_FILE) -> None:
    cutoff = (datetime.now() - timedelta(days=VERDICTS_MAX_AGE_DAYS)).strftime("%Y-%m-%d")
    kept = {k: v for k, v in verdicts.items() if isinstance(v, dict) and v.get("checked", "") >= cutoff}
    try:
        path.write_text(json.dumps(kept, ensure_ascii=False, indent=1, sort_keys=True), encoding="utf-8")
    except Exception as e:
        log.warning("Could not save %s: %s", path.name, e)

def cached_verdict(verdicts: dict, ann_id: str) -> dict | None:
    entry = verdicts.get(ann_id)
    if isinstance(entry, dict) and entry.get("v") == VERDICTS_VERSION:
        return entry
    return None

class ListingFetchError(RuntimeError):
    """A listing page could not be fetched after all retries."""

# ── Main Scraper Class ───────────────────────────────────────────────────────

class AlgerieMarchesScraper:

    def __init__(self):
        self.s = requests.Session()
        self.s.headers.update({
            "User-Agent": UA,
            "Accept": "application/json",
            "Origin": BASE_URL,
            "Referer": f"{BASE_URL}/login",
        })
        self._load_cookies()
        self.drive_mgr = GoogleDriveManager()
        self.failures: list[str] = []
        self.unverified: list[str] = []
        self.ad_errors = 0
        self._session_refreshed = False

    # ── Cookie persistence ───────────────────────────────────────────────────

    def _save_cookies(self):
        cookies = []
        for c in self.s.cookies:
            cookies.append({
                "name": c.name, "value": c.value,
                "domain": c.domain, "path": c.path,
                "secure": c.secure, "expires": c.expires,
            })
        with open(COOKIE_FILE, "w", encoding="utf-8") as f:
            json.dump(cookies, f, indent=2)
        log.info("Saved %d cookies to disk", len(cookies))

    def _load_cookies(self):
        if not COOKIE_FILE.exists():
            log.info("No saved cookies")
            return
        try:
            with open(COOKIE_FILE, "r", encoding="utf-8") as f:
                cookies = json.load(f)
            for c in cookies:
                self.s.cookies.set(
                    c["name"], c["value"],
                    domain=c.get("domain", ""),
                    path=c.get("path", "/"),
                    secure=c.get("secure", True),
                )
            log.info("Loaded %d cookies from disk", len(cookies))
        except Exception as e:
            log.warning("Failed to load cookies: %s", e)

    def _has_auth_cookies(self) -> bool:
        names = {c.name for c in self.s.cookies}
        return "__Secure-authjs.session-token" in names and "connect.sid" in names

    # ── Auth ─────────────────────────────────────────────────────────────────

    def _get_csrf(self) -> str:
        r = self.s.get(CSRF_URL, timeout=30)
        r.raise_for_status()
        return r.json()["csrfToken"]

    def _login(self, csrf: str) -> bool:
        payload = {
            "csrfToken": csrf,
            "identifier": EMAIL,
            "email": EMAIL,
            "password": PASSWORD,
            "callbackUrl": BASE_URL,
        }
        r = self.s.post(
            LOGIN_URL,
            data=payload,
            headers={
                "Content-Type": "application/x-www-form-urlencoded",
                "X-Auth-Return-Redirect": "1",
            },
            timeout=40,
        )
        has_auth = self._has_auth_cookies()
        try:
            res_url = r.json().get("url", "")
        except Exception:
            res_url = r.headers.get("Location", "")
        log.info("Login: status=%d result_url=%s has_auth=%s", r.status_code, res_url[:80], has_auth)
        return has_auth


    def _get_connect_sid_id(self) -> str:
        for c in self.s.cookies:
            if c.name == "connect.sid":
                raw = unquote(unquote(c.value))
                return raw.lstrip("s:").split(".")[0]
        return ""

    def _list_sessions(self) -> list[dict]:
        try:
            r = self.s.get(SESSIONS_URL, timeout=30)
            data = r.json()
            if isinstance(data, dict) and "sessions" in data:
                return data["sessions"]
            return data if isinstance(data, list) else []
        except Exception as e:
            log.warning("List sessions failed: %s", e)
            return []

    def _disconnect_one(self, session_id: str) -> dict:
        r = self.s.post(DISCONNECT_URL, json={"session_id": session_id}, timeout=30)
        try:
            return r.json()
        except Exception:
            return {"status": r.status_code, "text": r.text[:200]}

    def _promote_session(self) -> bool:
        try:
            csrf = self._get_csrf()
            r = self.s.post(
                f"{BASE_URL}/api/auth/session",
                json={"csrfToken": csrf, "data": {"sessionType": "normal"}},
                headers={"Content-Type": "application/json"},
                timeout=15,
            )
            data = r.json()
            is_normal = (
                data.get("sessionType") == "normal"
                or (data.get("user") or {}).get("sessionType") == "normal"
            )
            log.info("Session promotion: status=%d normal=%s", r.status_code, is_normal)
            return is_normal
        except Exception as e:
            log.warning("Session promotion failed: %s", e)
            return False

    def _test_detail_access(self) -> bool:
        if not self._has_auth_cookies():
            return False
        try:
            r = self.s.get(f"{API_BASE}/ads/all",
                           params={"type": "avis-attribution", "avis": "true", "page": 1, "limit": 1},
                           timeout=20)
            items = r.json().get("data", []) if isinstance(r.json(), dict) else []
            if not items:
                return False
            slug = items[0].get("slug", "")
            if not slug:
                return False

            r = self.s.get(f"{BASE_URL}/annonces/{slug}",
                           headers={"Accept": "text/html"},
                           timeout=30, allow_redirects=True)
            html = r.text
            if "manage-sessions" in html[:5000]:
                log.info("Detail test: redirected to manage-sessions (temporary)")
                return False
            if len(html) < 50000:
                log.info("Detail test: page too small (%d bytes)", len(html))
                return False
            log.info("Detail test: OK (%d bytes)", len(html))
            return True
        except Exception as e:
            log.warning("Detail test error: %s", e)
            return False

    def ensure_session(self):
        if not PASSWORD:
            raise RuntimeError("No password. Edit .env and set AM_PASSWORD=...")

        if self._has_auth_cookies():
            log.info("Have saved auth cookies, testing...")
            try:
                r_sess = self.s.get(f"{BASE_URL}/api/auth/session", timeout=15)
                sess_data = r_sess.json()
                if sess_data and sess_data.get("user"):
                    curr_type = sess_data.get("sessionType") or sess_data.get("user", {}).get("sessionType")
                    if curr_type == "temporary":
                        log.info("Saved session is temporary, promoting to normal...")
                        self._promote_session()
            except Exception as e:
                log.warning("Session check error: %s", e)

            if self._test_detail_access():
                log.info("Saved session works!")
                self._save_cookies()
                return
            log.info("Saved cookies didn't work, need fresh login")

        self.s.cookies.clear()
        csrf = self._get_csrf()
        if not self._login(csrf):
            log.warning("Login failed -- continuing without authenticated session (public mode)")
            return


        my_id = self._get_connect_sid_id()
        sessions = self._list_sessions()
        my_type = "?"
        competitors = []
        for s in sessions:
            sid = s.get("sessionId", "")
            if sid == my_id:
                my_type = s.get("type", "?")
            elif sid:
                competitors.append(sid)

        log.info("Sessions: %d total, my_id=%s type=%s, competitors=%d",
                 len(sessions), my_id[:10] if my_id else "?", my_type, len(competitors))

        if competitors and my_type == "temp":
            target = competitors[0]
            result = self._disconnect_one(target)
            log.info("Disconnected 1 competitor %s -> %s", target[:10],
                     json.dumps(result, ensure_ascii=False)[:200])
        else:
            log.info("No competitor disconnect needed (my_type=%s, competitors=%d)", my_type, len(competitors))

        self._promote_session()
        self._save_cookies()

    # ── Fetching ─────────────────────────────────────────────────────────────

    def fetch_listings(self, page: int = 1, notice_type: str = "avis-attribution") -> list[dict]:
        """Fetch listings from API with automatic fallback to proxy endpoint."""
        params = {"page": page, "limit": PAGE_SIZE}
        if notice_type == "appels-doffres":
            params["type"] = "appels-doffres"
        else:
            params["type"] = "avis-attribution"
            params["avis"] = "true"

        # Try API_BASE first, fallback to API_PROXY; retry the pair with backoff.
        # An empty list is only returned when the API really answered with no data.
        last_error = ""
        for attempt in range(LISTING_RETRIES):
            for endpoint in [f"{API_BASE}/ads/all", API_PROXY]:
                try:
                    r = self.s.get(endpoint, params=params, timeout=30)
                    if r.status_code == 200:
                        data = r.json()
                        if isinstance(data, dict) and isinstance(data.get("data"), list):
                            return data["data"]
                        last_error = f"{endpoint}: unexpected payload {str(data)[:80]}"
                    else:
                        last_error = f"{endpoint}: HTTP {r.status_code}"
                except Exception as e:
                    last_error = f"{endpoint}: {e}"
            log.warning("Listing page %d (%s) failed, attempt %d/%d: %s",
                        page, notice_type, attempt + 1, LISTING_RETRIES, last_error)
            if attempt + 1 < LISTING_RETRIES:
                time.sleep(LISTING_RETRY_DELAY * (attempt + 1))

        raise ListingFetchError(f"{notice_type} page {page}: {last_error}")

    def fetch_detail(self, slug: str) -> dict:
        url = f"{BASE_URL}/annonces/{slug}"
        r = self.s.get(url, headers={"Accept": "text/html"}, timeout=40, allow_redirects=True)
        if r.status_code != 200:
            return {"_error": f"HTTP {r.status_code}"}
        html = r.text

        if "manage-sessions" in html[:5000]:
            return {"_error": "SESSION_BLOCKED", "_html_len": len(html)}

        result = {"_source": "html", "_html_len": len(html), "_html": html}

        # Extract ad JSON from RSC payload
        unescaped = html.replace('\\"', '"').replace('\\\\', '\\')
        idx = unescaped.find('"ad":{')
        if idx >= 0:
            start = idx + 5
            depth, end = 0, start
            for i, ch in enumerate(unescaped[start:min(len(unescaped), start + 40000)]):
                if ch == '{': depth += 1
                elif ch == '}':
                    depth -= 1
                    if depth == 0:
                        end = start + i + 1
                        break
            try:
                ad = json.loads(unescaped[start:end])
                result["id_annonce"] = ad.get("id_annonce", "")
                result["titre_detail"] = ad.get("titre", "")
                result["date_parution"] = to_local_date(ad.get("date_parution"))
                result["meta_description"] = ad.get("meta_description", "")
                result["wilaya"] = (ad.get("wilaya") or {}).get("nom_fr", "")
                result["type_annonce"] = (ad.get("type_annonce") or {}).get("type_annonce_fr", "")
                result["sous_type"] = (ad.get("type_appel_offres") or {}).get("types_ao_fr", "")

                annonceurs = ad.get("annonceurs") or {}
                result["annonceur"] = annonceurs.get("annonceur") or ad.get("adresse_annonceur") or ""
                result["code_annonce"] = ad.get("code_annonce") or ""
                result["date_echeance"] = to_local_date(ad.get("date_echeance"))
                result["entreprise_concernee"] = ad.get("entreprise_concerne") or ""
                result["anep"] = ad.get("anep") or ""
                result["description"] = ad.get("description") or ""

            except Exception as e:
                log.warning("Failed to parse ad JSON for %s: %s", slug[:30], e)

        # Fallbacks from HTML text
        m = re.search(r'Annonceur</span>.*?capitalize[^>]*>([^<]+)<', unescaped, re.DOTALL)
        if m and m.group(1).strip() and not result.get("annonceur"):
            result["annonceur"] = m.group(1).strip()

        m = re.search(r'Code\s*N[^:]*:\s*(\S+)', unescaped)
        if m and not result.get("code_annonce"):
            result["code_annonce"] = m.group(1).strip()

        m = re.search(r"Date d.{1,5}ch[eé]ance.{0,300}?(\d{4}-\d{2}-\d{2})", unescaped, re.DOTALL)
        if m and not result.get("date_echeance"):
            result["date_echeance"] = m.group(1)

        m = re.search(r'Entreprise concern[eé]e.*?</span>\s*<[^>]*>([^<]+)<', unescaped, re.DOTALL | re.IGNORECASE)
        if m and m.group(1).strip() and not result.get("entreprise_concernee"):
            result["entreprise_concernee"] = m.group(1).strip()

        # Montant extraction: one amount, never digits glued across dates or lot numbers
        combined_text = f"{result.get('description', '')} {result.get('entreprise_concernee', '')}"
        montant = extract_amount(combined_text)
        if montant is not None:
            result["montant"] = montant

        return result

    def get_scan_urls(self, ad_item: dict, detail_html: str = "") -> list[str]:
        """Resolve all full URLs for attached newspaper scans, images, and PDFs."""
        urls = []

        # 1. Extract from detail page HTML regex (images, PDFs, attachments)
        if detail_html:
            # Match standard image/ad paths
            raw_urls = re.findall(r'/(?:api/(?:images|documents|files|attachments)/ads/[^\"\'\s>]+)', detail_html)
            for u in raw_urls:
                clean = u.rstrip('\\')
                full = f"{BASE_URL}{clean}" if clean.startswith("/") else clean
                if full not in urls:
                    urls.append(full)

            # Match any PDF links in the page
            pdf_matches = re.findall(r'(?:href|src)=[\"\']([^\"\']+\.pdf[^\"\'\s>]*)[\"\']', detail_html, re.IGNORECASE)
            for u in pdf_matches:
                clean = u.rstrip('\\').split("?")[0]
                if not clean.startswith("http"):
                    clean = f"{BASE_URL}{clean}" if clean.startswith("/") else f"{BASE_URL}/{clean}"
                if clean not in urls:
                    urls.append(clean)

        # 2. Extract from ad dictionary fields (image_principale, pdf, document, etc.)
        ann_id = str(ad_item.get("id_annonce", "")).strip()
        date_val = to_local_date(ad_item.get("date_parution"))
        try:
            dt = datetime.strptime(date_val, "%Y-%m-%d")
            y, m, d = dt.year, dt.month, dt.day
        except Exception:
            now = datetime.now()
            y, m, d = now.year, now.month, now.day

        keys = [
            "image_principale", "image_secondaire", "image_ternaire", "image_quaternaire", "image_5", "image_6",
            "pdf", "document", "fichier", "cahier_charges", "piece_jointe"
        ]
        for k in keys:
            val = ad_item.get(k)
            if val and isinstance(val, str):
                val_clean = val.strip()
                if val_clean.lower().endswith((".jpg", ".jpeg", ".png", ".pdf")):
                    if val_clean.startswith("http"):
                        if val_clean not in urls:
                            urls.append(val_clean)
                    else:
                        constructed = f"{BASE_URL}/api/images/ads/{y}/{m}/{d}/{ann_id}/{val_clean}"
                        if constructed not in urls:
                            urls.append(constructed)

        # Deduplicate URLs by file basename to prevent duplicate scan processing
        unique_urls = []
        seen_fnames = set()
        for u in urls:
            fn = u.split("/")[-1].split("?")[0].lower()
            if fn and fn not in seen_fnames:
                seen_fnames.add(fn)
                unique_urls.append(u)

        return unique_urls

    def collect_listings(self, notice_type: str, since_date: str = "") -> list[dict]:
        """All listings published on/after since_date, de-duplicated by notice ID."""
        label = "Appels d'Offres" if notice_type == "appels-doffres" else "Avis d'Attribution"
        all_annonces, seen = [], set()
        max_pages = MAX_PAGES_CATCHUP if since_date else MAX_PAGES
        reached_end = not since_date

        for page in range(1, max_pages + 1):
            try:
                listings = self.fetch_listings(page, notice_type=notice_type)
            except ListingFetchError as e:
                msg = f"{label} crawl stopped at page {page} -- {e}"
                log.error(msg)
                self.failures.append(msg)
                return all_annonces
            if not listings:
                log.info("  Page %d: 0 listings -- end of results", page)
                reached_end = True
                break

            page_dates = [d for d in (to_local_date(a.get("date_parution")) for a in listings) if d]
            log.info("  Page %d: %d listings (dates: %s to %s)",
                     page, len(listings), max(page_dates) if page_dates else "?", min(page_dates) if page_dates else "?")

            for a in listings:
                ann_id = str(a.get("id_annonce", "")).strip()
                if ann_id and ann_id in seen:
                    continue  # pages shift while new notices are published during the crawl
                d = to_local_date(a.get("date_parution"))
                if since_date and d and d < since_date:
                    continue
                if ann_id:
                    seen.add(ann_id)
                all_annonces.append(a)

            # Listings come newest first: stop once a whole page is older than the cutoff
            if since_date and page_dates and max(page_dates) < since_date:
                log.info("  Page %d is entirely older than cutoff %s -- stopping pagination", page, since_date)
                reached_end = True
                break
            time.sleep(0.4)

        if not reached_end:
            msg = (f"{label}: hit the {max_pages}-page cap before reaching {since_date} -- "
                   f"older notices in the window were not crawled")
            log.error(msg)
            self.failures.append(msg)
        return all_annonces

    def fetch_detail_with_retry(self, slug: str) -> dict:
        """Detail page data, re-authenticating once per run if the session got blocked; {} on failure."""
        if not slug:
            return {}
        detail = {}
        for attempt in range(3):
            try:
                detail = self.fetch_detail(slug)
            except Exception as e:
                log.warning("  -> Detail page %s failed (attempt %d): %s", slug[:40], attempt + 1, e)
                detail = {}
            if detail and not detail.get("_error"):
                return detail
            if detail.get("_error") == "SESSION_BLOCKED" and not self._session_refreshed:
                log.warning("  -> Session blocked on detail page -- re-authenticating once")
                self._session_refreshed = True
                try:
                    self.ensure_session()
                except Exception as e:
                    log.warning("  -> Re-authentication failed: %s", e)
                continue
            time.sleep(2 * (attempt + 1))
        log.warning("  -> Using listing data only for %s (%s)", slug[:40], detail.get("_error", "no detail"))
        return {}

    def download_scans(self, ad_item: dict, detail: dict, ann_id: str, target_date: str) -> list[Path]:
        scan_urls = self.get_scan_urls(ad_item, detail.get("_html", ""))
        if not scan_urls or not ann_id:
            return []
        ad_scan_dir = SCANS_DIR / target_date / ann_id
        ad_scan_dir.mkdir(parents=True, exist_ok=True)
        downloaded = []

        for u in scan_urls:
            fname = u.split("/")[-1].split("?")[0]
            local_fp = ad_scan_dir / fname
            if not local_fp.exists():
                for attempt in range(2):
                    try:
                        img_r = self.s.get(u, timeout=30)
                        if img_r.status_code == 200 and img_r.content:
                            local_fp.write_bytes(img_r.content)
                            log.info("  -> Downloaded scan: %s (%d bytes)", fname, len(img_r.content))
                            break
                        log.warning("  -> Scan %s returned HTTP %d", fname, img_r.status_code)
                        if img_r.status_code == 404:
                            break
                    except Exception as e:
                        log.warning("  -> Failed to download scan %s: %s", u, e)
                    time.sleep(2)

            if local_fp.exists() and local_fp not in downloaded:
                downloaded.append(local_fp)
                # Attempt Drive upload
                if self.drive_mgr:
                    self.drive_mgr.upload_scan(local_fp, target_date)

        if not downloaded:
            shutil.rmtree(ad_scan_dir, ignore_errors=True)
        return downloaded

    def read_scans(self, scans: list[Path], need_scope: bool, want_award_data: bool):
        """
        OCR the scan pages (up to MAX_SCAN_PAGES), reading further pages only while something is missing.
        Returns merged AI data, {} when every engine failed, or None when OCR was not attempted.
        """
        if not scans or not (NVIDIA_API_KEY or GEMINI_API_KEY):
            return None
        ai, any_read = {}, False
        for i, fp in enumerate(scans[:MAX_SCAN_PAGES]):
            if i > 0:
                missing_scope = need_scope and not ai_confirmed_scope(ai)
                missing_award = want_award_data and ai.get("budget", "/") == "/" and not ai.get("lots")
                if not (missing_scope or missing_award):
                    break
            engine = f"NVIDIA NIM ({NVIDIA_MODEL})" if NVIDIA_API_KEY and not _nvidia_state["disabled"] else "Gemini"
            log.info("  -> Scanning attached newspaper file (%s) with %s...", fp.name, engine)
            try:
                page = analyze_scan(fp.read_bytes())
            except Exception as e:
                log.warning("  -> AI Vision analysis error: %s", e)
                page = {}
            any_read = any_read or bool(page)
            ai = merge_ai_data(ai, page)
        return ai if any_read else {}

    def process_ad(self, a: dict, notice_type: str, run_id: str, decision: str, text_scope: str | None,
                   verdicts: dict) -> dict | None:
        """Fetch, OCR and classify one notice. Returns the result row, or None when it is not for TAPIDOR."""
        today_str = datetime.now().strftime("%Y-%m-%d")
        ann_id = str(a.get("id_annonce", "")).strip()
        slug = a.get("slug", "")
        titre = a.get("titre", "") or ""
        raw_date = to_local_date(a.get("date_parution")) or today_str

        detail = self.fetch_detail_with_retry(slug)
        target_date = detail.get("date_parution") or raw_date

        # ── 1. Download Attached Scans & Upload to Google Drive ──
        scans = self.download_scans(a, detail, ann_id, target_date)

        # ── 2. Scope gate: explicit text wins; otherwise the scan must name gazon or tapis ──
        combined_text = f"{titre} {detail.get('description', '')} {a.get('description', '')}"
        scope = text_scope or detect_scope(combined_text)
        ai = self.read_scans(scans, need_scope=not scope, want_award_data=(notice_type == "avis-attribution"))

        def reject(reason: str):
            log.info("  -> [REJECTED] %s (%s) -- Skipping.", titre[:65], reason)
            if scans:
                shutil.rmtree(SCANS_DIR / target_date / ann_id, ignore_errors=True)
            if ann_id:
                verdicts[ann_id] = {"verdict": "reject", "v": VERDICTS_VERSION, "checked": today_str,
                                    "type": notice_type, "reason": reason}

        if not scope:
            scope = ai_confirmed_scope(ai or {})
            if not scope:
                if scans and not ai:
                    # OCR unavailable (rate limits, outage, no key): do not drop the notice for good
                    log.warning("  -> [UNVERIFIED] %s -- scan could not be read; will retry next run", titre[:65])
                    self.unverified.append(ann_id or slug)
                    return None
                reject("no explicit gazon/tapis in text or scan" if scans else "no explicit gazon/tapis and no scan")
                return None
        ai = ai or {}

        # ── 3. Lots: keep the in-scope ones, each becomes its own Sheet row for awards ──
        lots = ai.get("lots") or []
        for lot in lots:
            lot["in_scope"] = bool(lot.get("in_scope") or detect_scope(lot.get("objet", "")))
        if len(lots) > 1:
            lots = [lot for lot in lots if lot["in_scope"]] or lots

        # ── 4. Merge Web Data with AI Complementary Data ────────
        wilaya_val = (
            detail.get("wilaya")
            or (a.get("wilaya") or {}).get("nom_fr", "")
            or (ai.get("wilaya", "/").upper() if ai.get("wilaya", "/") != "/" else "")
        )
        annonceur_val = detail.get("annonceur") or (ai.get("annonceur") if ai.get("annonceur", "/") != "/" else "")

        commune_val = parse_commune(titre, annonceur_val)
        if (commune_val == "/" or len(commune_val) <= 2) and ai.get("commune", "/") != "/":
            commune_val = strip_accents(ai["commune"]).upper().strip()

        entreprise_val = detail.get("entreprise_concernee", "")
        if not entreprise_val and ai.get("entreprise_attributaire", "/") != "/":
            entreprise_val = ai["entreprise_attributaire"]

        montant_val = detail.get("montant", "")
        if montant_val in ("", None) and ai.get("budget", "/") != "/":
            montant_val = ai["budget"]

        delai_val = parse_delai(detail.get("nbr_jours", 0), detail.get("description", ""))
        if delai_val == "/" and ai.get("delai", "/") != "/":
            delai_val = normalize_delai(ai["delai"])

        action_val = classify_action(titre)
        if action_val == "TRAVAUX" and ai.get("action", "/") != "/":
            action_val = classify_action(ai["action"])

        ptype_val = classify_type_projet(titre)
        if ptype_val in GENERIC_PTYPES:
            in_scope_objets = [lot.get("objet", "/") for lot in lots if lot.get("in_scope")]
            for text in in_scope_objets + [ai.get("type_projet", "/"), ai.get("objet_marche", "/")]:
                if text and text != "/" and classify_type_projet(text) not in GENERIC_PTYPES:
                    ptype_val = classify_type_projet(text)
                    break

        row = {
            "run_id": run_id,
            "id": ann_id,
            "slug": slug,
            "url": f"{BASE_URL}/annonces/{slug}",
            "titre": titre,
            "scope": scope,
            "action": action_val,
            "ptype": ptype_val,
            "wilaya": wilaya_val,
            "commune": commune_val,
            "date_parution": detail.get("date_parution") or raw_date,
            "date_parution_ao": ai.get("date_parution_ao", "/"),
            "date_echeance": detail.get("date_echeance", ""),
            "type_annonce": notice_type,
            "annonceur": annonceur_val,
            "code_annonce": detail.get("code_annonce", ""),
            "entreprise_concernee": entreprise_val,
            "montant": montant_val,
            "delai": delai_val,
            "lots": lots,
            "nbr_jours": detail.get("nbr_jours", 0),
            "scans": [str(p) for p in scans],
            "scraped_at": datetime.now().isoformat(),
        }
        if ann_id:
            verdicts[ann_id] = {"verdict": "match", "v": VERDICTS_VERSION, "checked": today_str,
                                "type": notice_type, "row": row}
        log.info("  [MATCH %s] %s | %s | Wilaya: %s | Commune: %s | Lots: %d | Annonceur: %s",
                 scope, ann_id, titre[:35], row["wilaya"], row["commune"], len(lots), row["annonceur"][:25])
        return row

    def scrape_notice_type(self, notice_type: str, run_id: str, since_date: str = "", verdicts: dict | None = None) -> list[dict]:
        """Scrape either 'appels-doffres' or 'avis-attribution' with scan downloads and AI extraction."""
        verdicts = verdicts if verdicts is not None else {}
        label = "Appels d'Offres" if notice_type == "appels-doffres" else "Avis d'Attribution"
        log.info("--- Scraping %s (since date: %s) ---", label, since_date or "any")

        all_annonces = self.collect_listings(notice_type, since_date)
        log.info("Total %s fetched (>= %s): %d", label, since_date or "all", len(all_annonces))

        candidates = []
        n_cached_rejects = 0
        for a in all_annonces:
            ann_id = str(a.get("id_annonce", "")).strip()
            entry = cached_verdict(verdicts, ann_id) if ann_id else None
            if entry and entry.get("verdict") == "reject":
                n_cached_rejects += 1
                continue
            if entry and entry.get("verdict") == "match" and isinstance(entry.get("row"), dict):
                candidates.append((a, "cached", entry["row"]))
                continue

            t = str(a.get("titre") or "")
            decision, info = prefilter_decision(f"{t} {a.get('description') or ''}")
            if decision == "skip":
                if info.startswith("excluded"):
                    log.info("  -> Pre-filter excluded notice: %s [%s]", t[:70], info)
                continue
            candidates.append((a, decision, info))

        counts = {k: sum(1 for _, d, _ in candidates if d == k) for k in ("explicit", "ocr", "cached")}
        log.info("After TAPIDOR scope filter: %d %s to process (%d explicit, %d to verify by OCR, "
                 "%d cached matches; %d cached rejections skipped)",
                 len(candidates), label, counts["explicit"], counts["ocr"], counts["cached"], n_cached_rejects)

        results = []
        for a, decision, info in candidates:
            if decision == "cached":
                log.info("  [CACHED MATCH] %s | %s", info.get("id"), str(info.get("titre", ""))[:50])
                results.append(info)
                continue
            try:
                row = self.process_ad(a, notice_type, run_id, decision,
                                      info if decision == "explicit" else None, verdicts)
            except Exception as e:
                # One malformed notice must never stop the whole cycle
                log.exception("  -> Error while processing notice %s (%s): %s",
                              a.get("id_annonce"), str(a.get("titre", ""))[:50], e)
                self.ad_errors += 1
                continue
            if row:
                results.append(row)
            time.sleep(0.5)

        return results

    # ── Main Run ─────────────────────────────────────────────────────────────

    def run(self, since_date: str | None = None, notice_types: list[str] | None = None) -> int:
        """Run the daily cycles. Returns the process exit code (1 when something needs attention)."""
        run_id = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        if not since_date:
            since_date = os.getenv("AM_SINCE_DATE", "").strip()
        if not since_date:
            # Default to last 7 days so weekend publications and downtime gaps are automatically recovered
            since_date = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d")

        self.failures, self.unverified, self.ad_errors = [], [], 0
        self._session_refreshed = False
        _nvidia_state.update(failures=0, disabled=False)
        verdicts = load_verdicts()

        log.info("=" * 60)
        log.info("AlgerieMarches Scraper Run %s started", run_id)
        log.info("Cutoff date (since): %s", since_date)
        log.info("Keyword filter: %s", ACTIVE_KEYWORDS)

        cycles = [
            ("appels-doffres", "APPELS D'OFFRES", "AO"),    # -> APPEL D'OFFRE 2026
            ("avis-attribution", "AVIS D'ATTRIBUTION", "AA"),  # -> AVIS D'ATTRIBUTIONS 2026
        ]
        all_matches = []
        try:
            self.ensure_session()
            for n, (notice_type, title, short) in enumerate(cycles, start=1):
                if notice_types and notice_type not in notice_types:
                    continue
                log.info("\n>>> CYCLE %d: %s <<<", n, title)
                try:
                    results = self.scrape_notice_type(notice_type, run_id, since_date=since_date, verdicts=verdicts)
                except Exception as e:
                    log.exception("%s cycle crashed: %s", title, e)
                    self.failures.append(f"{title} cycle crashed: {e}")
                    continue
                all_matches.extend(results)
                if results:
                    added = append_to_gsheet(results, notice_type=notice_type)
                    if added < 0:
                        self.failures.append(f"{title}: Google Sheet sync failed")
                    else:
                        log.info("%s Sync summary: %d added to Google Sheet", short, added)

            # Generate daily scans gallery for GitHub
            if all_matches:
                generate_daily_scans_summary(all_matches)
        finally:
            save_verdicts(verdicts)
            self._save_cookies()

        if self.unverified:
            log.warning("%d notice(s) could not be verified because no OCR engine answered: %s",
                        len(self.unverified), ", ".join(self.unverified))
        if self.ad_errors:
            self.failures.append(f"{self.ad_errors} notice(s) failed with an exception (see traceback above)")
        if self.failures:
            for problem in self.failures:
                log.error("RUN PROBLEM: %s", problem)
            log.error("Run %s finished with %d problem(s)", run_id, len(self.failures))
            return 1
        log.info("\nRun %s successfully completed!", run_id)
        return 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="AlgerieMarches Scraper & Google Sheets Sync")
    parser.add_argument("--since", dest="since_date", default=None,
                        help="Cutoff publication date YYYY-MM-DD (defaults to 7 days ago)")
    parser.add_argument("--type", dest="notice_type", default=None,
                        choices=["appels-doffres", "avis-attribution"],
                        help="Only scrape this notice type (default: both)")
    args = parser.parse_args()

    scraper = AlgerieMarchesScraper()
    sys.exit(scraper.run(since_date=args.since_date,
                         notice_types=[args.notice_type] if args.notice_type else None))
