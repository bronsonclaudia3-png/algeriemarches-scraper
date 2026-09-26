"""
AlgerieMarches Scraper v4 — Dual Scrape (Appels d'Offres & Avis d'Attributions)
Automated daily sync to Google Sheets & local Downloads Excel file.

Targets:
  1. "APPEL D'OFFRE 2026" (11 columns)
  2. "AVIS D'ATTRIBUTIONS 2026" (15 columns)
Domain:
  Artificial turf, stadiums, sports facilities, playgrounds, school courtyards (TAPIDOR).
"""

import argparse
import base64
import csv
import json
import logging
import os
import re
import sys
import time
import unicodedata
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import unquote

import gspread
import requests
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload

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
NVIDIA_API_KEY = os.getenv("NVIDIA_API_KEY", "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").strip()
GOOGLE_DRIVE_FOLDER_ID = os.getenv("GOOGLE_DRIVE_FOLDER_ID", "").strip()

BASE_URL = "https://algeriemarches.com"
API_BASE = "https://api.algeriemarches.com/api"
API_PROXY = f"{BASE_URL}/api/proxy/ads"

CSRF_URL = f"{BASE_URL}/api/auth/csrf"
LOGIN_URL = f"{BASE_URL}/api/auth/callback/credentials"
SESSIONS_URL = f"{BASE_URL}/api/auth/active-sessions"
DISCONNECT_URL = f"{BASE_URL}/api/proxy/auth/disconnect-session"

PAGE_SIZE = 20
MAX_PAGES = 3

# Sport & turf keywords for TAPIDOR (Exclusive Gazon Scope)
DEFAULT_KEYWORDS = r"\b(?:gazon|pelouse|engazonnement|عشب|تعشيب|نجيل|نجيلة|stade|terrain|football|sport)\b"
KEYWORDS_ENV = os.getenv("AM_KEYWORDS", "").strip()
ACTIVE_KEYWORDS = KEYWORDS_ENV if KEYWORDS_ENV else DEFAULT_KEYWORDS
KEYWORD_FILTER = re.compile(ACTIVE_KEYWORDS, re.IGNORECASE)

GAZON_EXPLICIT_KEYWORDS = re.compile(
    r"\b(?:gazon|pelouse|engazonnement|عشب|تعشيب|نجيل|نجيلة|gazonné|gazonne|gazonnee)\b",
    re.IGNORECASE
)

NON_GAZON_EXCLUSIONS = re.compile(
    r"\b(?:cour\s+de\s+justice|cours\s+de\s+justice|tribunal|tribunaux|مجلس\s+قضاء|محكمة|"
    r"chauffage|climatisation|plomberie|electricite|électricité|éclairage\s+public|eclairage\s+public|"
    r"mobilier|bureau|fourniture\s+de\s+bureau|fournitures\s+de\s+bureau|عتاد\s+مكتبي|أثاث|اثاث|طباعة|"
    r"assiette\s+de\s+terrain|assiette\s+fonciere|assiette\s+foncière|logement|logements|aadl|lpp|lpa|pos|viabilisation|geotechnique|géotechnique|"
    r"assainissement|assainisement|aep|drainage|cantine|salle\s+omnisport|salle\s+de\s+sport|piscine)\b",
    re.IGNORECASE
)

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
)

# ── Classification & Parsing Helpers ─────────────────────────────────────────

def strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")

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

    facility_patterns = [
        ("STADE DE PROXIMITÉ", r"\bSTADES?\s+(?:DE\s+)?PROXIMIT[EÉ]\b"),
        ("STADE DE FOOTBALL", r"\bSTADES?\s+DE\s+FOOTBALL\b|\bSTADES?\s+DE\s+FOOT\b"),
        ("STADE COMMUNAL", r"\bSTADES?\s+COMMUNAU?X?\b|\bSTADES?\s+COMMUNALES?\b"),
        ("STADE MUNICIPAL", r"\bSTADES?\s+MUNICIPAU?X?\b|\bSTADES?\s+MUNICIPALES?\b"),
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

def parse_commune(titre: str, annonceur: str) -> str:
    combined = f"{annonceur} {titre}".upper()
    m = re.search(r"\b(?:COMMUNE D'|COMMUNE DE |APC D'|APC DE )([A-Z\s'\-]+?)(?:\bWILAYA|\bDAIRA|\bALIMENTATION|$)", combined)
    if m:
        c = m.group(1).strip()
        return c if len(c) < 30 else c[:30].strip()
    if "COMMUNE" in annonceur.upper():
        words = annonceur.upper().replace("COMMUNE", "").replace("D'", "").replace("DE", "").strip()
        if words:
            return words[:30].strip()
    return "/"

def clean_annonceur(ann_raw: str) -> str:
    if not ann_raw:
        return "/"
    ann_upper = ann_raw.upper().strip()
    if "COMMUNE" in ann_upper or "A.P.C" in ann_upper or "APC" in ann_upper:
        return "COMMUNE"
    if "DJS" in ann_upper or "JEUNESSE ET SPORT" in ann_upper or "JEUNESSE ET DES SPORT" in ann_upper:
        return "DJS DE LA WILAYA"
    if "DEP" in ann_upper or "EQUIPEMENT PUBLIC" in ann_upper or "EQUIPEMENTS PUBLICS" in ann_upper:
        return "DEP DE LA WILAYA"
    if "ADMINISTRATION LOCALE" in ann_upper or "D.A.L" in ann_upper or "DAL" in ann_upper:
        return "DIRECTION DE L'ADMINISTRATION LOCALE DE LA WILAYA"
    if "EDUCATION" in ann_upper:
        return "DIRECTION DE L'EDUCATION DE LA WILAYA"
    if "FORMATION ET DE L'ENSEIGNEMENT" in ann_upper:
        return "DIRECTION DE LA FORMATION ET DE L'ENSEIGNEMENT PROFESSIONNELS"
    return ann_raw[:50].strip()

def parse_budget(montant_str):
    if not montant_str or str(montant_str).strip() == "/":
        return "/"
    s = str(montant_str).upper().replace("DA", "").replace("DZD", "").strip()
    s = s.replace(" ", "")
    if "." in s and "," in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s and "." not in s:
        parts = s.split(",")
        if len(parts) == 2 and len(parts[1]) <= 2:
            s = f"{parts[0]}.{parts[1]}"
        else:
            s = s.replace(",", "")
    elif "." in s and "," not in s:
        parts = s.split(".")
        if len(parts) == 2 and len(parts[1]) <= 2:
            pass
        elif len(parts) > 2:
            if len(parts[-1]) <= 2:
                s = f"{''.join(parts[:-1])}.{parts[-1]}"
            else:
                s = "".join(parts)
    try:
        val = float(re.sub(r"[^\d.]", "", s))
        return int(val) if val.is_integer() else val
    except Exception:
        return str(montant_str).strip()

def parse_delai(nbr_jours: int, description: str) -> str:
    if nbr_jours and int(nbr_jours) > 0:
        return f"{nbr_jours} JOURS"
    m = re.search(r"\b(\d+\s*(?:JOURS?|MOIS|SEMAINES?))\b", str(description).upper())
    if m:
        return m.group(1).strip()
    return "/"

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

NVIDIA_MODEL = "deepseek-ai/deepseek-v4.1-flash"
NVIDIA_FALLBACK_MODELS = [
    ("deepseek-ai/deepseek-v4.1-flash", 20),
    ("meta/llama-3.2-11b-vision-instruct", 25),
    ("moonshotai/kimi-k3", 35),
]

PROMPT_SCAN_ANALYSIS = """You are an expert document analysis and OCR system specialized in Algerian public procurement (Marchés Publics / الصفقات العمومية) for TAPIDOR, a company specializing exclusively in GAZON (artificial turf / pelouse synthétique / engazonnement / sports turf).
The document may be in French, Arabic, or bilingual.

CRITICAL REQUIREMENT (GAZON SCOPE FILTER):
Determine whether this project specifically involves GAZON (artificial turf / pelouse synthétique / engazonnement / natural sports turf).
If the project is NOT about gazon (for example: it is civil works, building construction, concrete or asphalt paving without turf, indoor sports halls without turf, heating, office furniture, geotechnical study, housing, sanitation/sewerage, electrical works):
Set "is_gazon": false.
Only set "is_gazon": true if the project explicitly involves GAZON / PELOUSE / ENGAZONNEMENT (e.g. fourniture et pose de gazon synthétique, engazonnement de terrain/stade, revêtement en gazon synthétique / تغطية بالعشب الاصطناعي / تعشيب).

DATE DE PARUTION REQUIREMENT (FOR AVIS D'ATTRIBUTION):
Extract "date_parution_ao" strictly from INSIDE the text paragraph (where it mentions when the original Appel d'Offres was published, e.g. 'paru le 22/07/2026' or 'الصادرة بتاريخ 2026/07/22').
If NO publication date is mentioned in the paragraph text, return '/'.
NEVER use the newspaper print date or footer date at the very bottom (such as 'An-Nasr 21-9-2026' or 'El-Moudjahid 22-09-2026').

Extract the following in strict JSON:
{
  "is_gazon": true or false,
  "action": "Main action verb in French (e.g. REALISATION, AMENAGEMENT, REVETEMENT, REHABILITATION, REFECTION) or '/'",
  "type_projet": "Facility type in French (e.g. STADE, STADE COMMUNAL, STADE DE PROXIMITE, TERRAIN DE SPORT, GAZON SYNTHETIQUE) or '/'",
  "budget": "Final winning amount or estimated budget with currency. Prioritize final corrected TTC amount. Example: '37 756 320,00 DA' or '/' if none.",
  "delai": "Execution delay. Format 'XX JOURS' or 'XX MOIS'. Example: '03 MOIS' or '60 JOURS' or '/' if none.",
  "commune": "Commune (municipality) name in UPPERCASE Latin/French letters or '/'",
  "wilaya": "Wilaya name in UPPERCASE or '/'",
  "annonceur": "Contracting authority name in French or '/'",
  "entreprise_attributaire": "Winning contractor name or '/' if none.",
  "date_parution_ao": "Date when original Appel d'Offres was published strictly from INSIDE the paragraph text (or '/' if not in paragraph text)."
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
    Extract complementary details from attached ad scans using NVIDIA NIM free endpoints.
    Primary model: deepseek-ai/deepseek-v4.1-flash.
    """
    key = api_key or NVIDIA_API_KEY
    if not key or not img_bytes:
        return {}

    opt_bytes, mime = prepare_image_for_ocr(img_bytes)
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
            "max_tokens": 1024,
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
                    log.info("NVIDIA NIM (%s) extracted -> is_gazon: %s | Commune: %s | Budget: %s | Delai: %s | Parution AO: %s",
                             model, data.get("is_gazon"), data.get("commune"), data.get("budget"), data.get("delai"), data.get("date_parution_ao"))
                    time.sleep(0.5)
                    return data
            elif r.status_code in (429, 503):
                log.warning("NVIDIA NIM (%s) status %d -- falling back...", model, r.status_code)
                continue
            else:
                log.warning("NVIDIA NIM (%s) returned %d: %s", model, r.status_code, r.text[:120])
        except Exception as e:
            log.warning("NVIDIA NIM (%s) request error (%s) -- falling back...", model, e)

    return {}

GEMINI_MODELS = [
    "gemini-3.5-flash-lite",
    "gemini-flash-latest",
    "gemini-3.5-flash",
    "gemini-3.6-flash",
    "gemini-2.5-flash",
]

def analyze_scan_with_gemini(img_bytes: bytes, api_key: str | None = None) -> dict:
    key = api_key or GEMINI_API_KEY
    if not key or not img_bytes:
        return {}

    b64_img = base64.b64encode(img_bytes).decode("utf-8")
    if img_bytes[:4] == b"%PDF":
        mime = "application/pdf"
    elif img_bytes[:8] == b"\x89PNG\r\n\x1a\n":
        mime = "image/png"
    else:
        mime = "image/jpeg"

    payload = {
        "contents": [
            {
                "parts": [
                    {"text": PROMPT_SCAN_ANALYSIS},
                    {"inline_data": {"mime_type": mime, "data": b64_img}}
                ]
            }
        ],
        "generationConfig": {"response_mime_type": "application/json"}
    }

    for model in GEMINI_MODELS:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={key}"
        try:
            r = requests.post(url, json=payload, headers={"Content-Type": "application/json"}, timeout=15)
            if r.status_code == 200:
                content = r.json()["candidates"][0]["content"]["parts"][0]["text"]
                data = json.loads(content)
                log.info("Gemini Vision (%s) extracted -> is_gazon: %s | Commune: %s | Budget: %s | Delai: %s | Parution AO: %s",
                         model, data.get("is_gazon"), data.get("commune"), data.get("budget"), data.get("delai"), data.get("date_parution_ao"))
                time.sleep(0.5)
                return data
            elif r.status_code in (429, 503):
                log.warning("Gemini (%s) status %d -- falling back immediately...", model, r.status_code)
                continue
            else:
                log.warning("Gemini (%s) returned %d: %s", model, r.status_code, r.text[:120])
        except Exception as e:
            log.warning("Gemini (%s) request error (%s) -- falling back...", model, e)
    return {}


def analyze_scan(img_bytes: bytes) -> dict:
    """
    Primary scan analysis using NVIDIA NIM (DeepSeek v4.1-flash).
    Falls back to Gemini if NVIDIA key is absent or NIM endpoints fail.
    """
    if NVIDIA_API_KEY:
        try:
            data = analyze_scan_with_nvidia(img_bytes)
            if data:
                return data
        except Exception as e:
            log.warning("NVIDIA NIM analysis error: %s", e)

    if GEMINI_API_KEY:
        try:
            return analyze_scan_with_gemini(img_bytes)
        except Exception as e:
            log.warning("Gemini fallback analysis error: %s", e)

    return {}

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

def append_to_gsheet(results: list[dict], notice_type: str, sheet_id: str | None = None) -> int:
    """
    Append results to the matching worksheet in Google Sheet.
    - notice_type 'appels-doffres' -> 'APPEL D'OFFRE 2026' (11 columns)
    - notice_type 'avis-attribution' -> 'AVIS D'ATTRIBUTIONS 2026' (15 columns)
    """
    sheet_id = sheet_id or os.getenv("GOOGLE_SHEET_ID", "1y5-OxNeL_zBhCUNNEVoh1nKZVsyBYrq5EJcgKBEt920")
    if not sheet_id:
        log.info("No GOOGLE_SHEET_ID configured -- skipping Google Sheet sync.")
        return 0

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
    elif sa_file.exists():
        try:
            gc = gspread.service_account(filename=str(sa_file))
            log.info("Authenticated with Google via state/service_account.json")
        except Exception as e:
            log.error("Failed to authenticate with service_account.json: %s", e)

    if not gc:
        log.info("No Google service account credentials found -- skipping Google Sheet sync.")
        return 0

    try:
        sh = gc.open_by_key(sheet_id)
        log.info("Opened Google Sheet: %s", sh.title)
    except Exception as e:
        log.warning("Could not access Google Sheet '%s': %s", sheet_id, e)
        log.warning(">> Please ensure you clicked 'Share' on the Google Sheet and added:")
        log.warning(">> algeriemarches-bot@project-df86d806-9e9d-42be-a7d.iam.gserviceaccount.com as Editor!")
        return 0

    target_title = "APPEL D'OFFRE 2026" if notice_type == "appels-doffres" else "AVIS D'ATTRIBUTIONS 2026"
    target_ws = None
    for ws in sh.worksheets():
        if ws.title.strip().upper() == target_title.upper():
            target_ws = ws
            break

    if not target_ws:
        log.warning("Worksheet '%s' not found in Google Sheet", target_title)
        return 0

    log.info("Target Google Sheet worksheet: %s", target_ws.title)

    try:
        header_col = "L3" if notice_type == "appels-doffres" else "P3"
        header_val = target_ws.acell(header_col).value
        if header_val != "SCANS":
            target_ws.update(header_col, [["SCANS"]])
            log.info("Added SCANS column header to %s in %s", target_ws.title, header_col)
    except Exception as e:
        log.debug("Header check notice: %s", e)

    try:
        all_values = target_ws.get_all_values()
    except Exception as e:
        log.error("Failed to read values from Google Sheet: %s", e)
        return 0

    last_num = 0
    existing_items = set()

    for idx, row in enumerate(all_values):
        if idx < 3: # Skip title & header rows
            continue
        if row and len(row) > 0 and str(row[0]).strip():
            try:
                n = int(str(row[0]).strip())
                if n > last_num:
                    last_num = n
            except Exception:
                pass
            if notice_type == "appels-doffres":
                action = row[2].strip().upper() if len(row) > 2 else ""
                t_proj = row[4].strip().upper() if len(row) > 4 else ""
                wilaya = row[8].strip().upper() if len(row) > 8 else ""
                commune = row[9].strip().upper() if len(row) > 9 else ""
                existing_items.add((action, t_proj, wilaya, commune))
            else:
                attr_par = row[12].strip().upper() if len(row) > 12 else ""
                wilaya = row[8].strip().upper() if len(row) > 8 else ""
                t_proj = row[4].strip().upper() if len(row) > 4 else ""
                action = row[2].strip().upper() if len(row) > 2 else ""
                commune = row[9].strip().upper() if len(row) > 9 else ""
                if attr_par:
                    existing_items.add((attr_par, wilaya, t_proj))
                if action and t_proj and wilaya:
                    existing_items.add(("ALT", action, t_proj, wilaya, commune))

    log.info("Google Sheet '%s' has %d rows (last N°: %d)", target_ws.title, len(all_values) - 3, last_num)

    rows_to_append = []
    today_str = datetime.now().strftime("%Y-%m-%d")

    for item in results:
        titre = item.get("titre", "")
        annonceur = item.get("annonceur", "")
        action = item.get("action") or classify_action(titre)
        ptype = item.get("ptype") or classify_type_projet(titre)
        wilaya = (item.get("wilaya") or "").strip().upper()
        commune = item.get("commune") if (item.get("commune") and item.get("commune") != "/") else parse_commune(titre, annonceur)
        ann_val = clean_annonceur(annonceur)

        dt_parution = str(item.get("date_parution", ""))[:10] or "/"
        dt_echeance = str(item.get("date_echeance", ""))[:10] or "/"

        ann_id = str(item.get("id", "") or item.get("id_annonce", "")).strip()
        scans_list = item.get("scans", [])
        if scans_list and ann_id and dt_parution != "/":
            scan_names = [Path(s).name for s in scans_list]
            repo_base = "https://github.com/bronsonclaudia3-png/algeriemarches-scraper"
            if len(scan_names) == 1:
                scan_link = f'=HYPERLINK("{repo_base}/blob/main/data/scans/{dt_parution}/{ann_id}/{scan_names[0]}", "📄 Voir Scan")'
            else:
                scan_link = f'=HYPERLINK("{repo_base}/tree/main/data/scans/{dt_parution}/{ann_id}", "📁 Voir {len(scan_names)} Scans")'
        else:
            scan_link = "/"

        if notice_type == "appels-doffres":
            dedup_key = (action.upper(), ptype.upper(), wilaya, commune.upper())
            if dedup_key in existing_items:
                log.info("  Already in Google Sheet (AO): %s | %s (%s)", action, ptype, wilaya)
                continue

            last_num += 1
            existing_items.add(dedup_key)

            date_col_val = dt_parution if dt_parution != "/" else today_str

            row_data = [
                last_num,                                # Col 1: N°
                date_col_val,                            # Col 2: DATE (aligns with publication date)
                action,                                  # Col 3: TITRE D'APPEL D'OFFRE
                1,                                       # Col 4: Nombre de projet
                ptype,                                   # Col 5: TYPE DE PROJET
                dt_parution,                             # Col 6: DATE DE PARUTION
                dt_echeance,                             # Col 7: DATE D'ECHEANCE
                ann_val,                                 # Col 8: ANNONCEUR
                wilaya if wilaya else "/",               # Col 9: WILAYA
                commune,                                 # Col 10: COMMUNE
                "TRAVAUX PUBLICS",                       # Col 11: CATEGORIE
                scan_link,                               # Col 12: SCANS
            ]
            rows_to_append.append(row_data)
        else:
            attr_par = item.get("entreprise_concernee", "").strip()
            dedup_key = (attr_par.upper(), wilaya, ptype.upper())
            dedup_key_alt = ("ALT", action.upper(), ptype.upper(), wilaya, commune.upper())
            if (attr_par != "" and dedup_key in existing_items) or (dedup_key_alt in existing_items):
                log.info("  Already in Google Sheet (AA): %s (%s)", attr_par[:30], wilaya)
                continue

            last_num += 1
            if attr_par:
                existing_items.add(dedup_key)
            existing_items.add(dedup_key_alt)

            budget = parse_budget(item.get("montant", "")) if item.get("montant") else "/"
            delai = item.get("delai") or parse_delai(item.get("nbr_jours", 0), item.get("description", ""))

            date_col_val = dt_parution if dt_parution != "/" else today_str
            # Col 6: Must strictly be the original AO publication date from INSIDE the text paragraph. Never the bottom date.
            dt_parution_ao = str(item.get("date_parution_ao", "")).strip() or "/"

            row_data = [
                last_num,                                # Col 1: N°
                date_col_val,                            # Col 2: DATE (aligns with publication date)
                action,                                  # Col 3: TITRE D'AVIS D'ATTRIBUTION
                1,                                       # Col 4: Nombre de projet
                ptype,                                   # Col 5: TYPE DE PROJET
                dt_parution_ao,                          # Col 6: DATE DE PARUTION (from inside paragraph or '/')
                dt_echeance,                             # Col 7: DATE D'ECHEANCE
                ann_val,                                 # Col 8: ANNONCEUR
                wilaya if wilaya else "/",               # Col 9: WILAYA
                commune,                                 # Col 10: COMMUNE
                dt_parution,                             # Col 11: DATE D'ATTRIBUTION
                budget,                                  # Col 12: BUDGET
                attr_par if attr_par else "/",           # Col 13: ATTRIBUTION PAR
                delai,                                   # Col 14: DELAI
                "/",                                     # Col 15: WILAYA 2
                scan_link,                               # Col 16: SCANS
            ]
            rows_to_append.append(row_data)

    if rows_to_append:
        try:
            target_ws.append_rows(rows_to_append, value_input_option="USER_ENTERED")
            log.info("Successfully appended %d new rows to Google Sheet '%s'", len(rows_to_append), target_ws.title)
        except Exception as e:
            log.error("Failed to append rows to Google Sheet: %s", e)
            return 0
    else:
        log.info("No new rows needed for Google Sheet '%s' (all up to date)", target_ws.title)

    return len(rows_to_append)

# ── Daily Scans Summary Generator ────────────────────────────────────────────

def generate_daily_scans_summary(all_results: list[dict]) -> list[Path]:
    """Generates clean visual Markdown galleries in data/scans/YYYY-MM-DD/README.md for GitHub."""
    by_date: dict[str, list[dict]] = {}
    today_str = datetime.now().strftime("%Y-%m-%d")
    for r in all_results:
        d = str(r.get("date_parution") or today_str)[:10]
        by_date.setdefault(d, []).append(r)

    generated = []
    for d_str, matches in by_date.items():
        day_dir = SCANS_DIR / d_str
        day_dir.mkdir(parents=True, exist_ok=True)
        md_file = day_dir / "README.md"
        lines = [
            f"# 📰 Newspaper Scans Gallery — {d_str}\n",
            "> Scans downloaded & OCR-analyzed with Gemini 2.5 Flash for TAPIDOR\n",
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
            "> Central archive of newspaper scans downloaded and OCR-analyzed by Gemini 2.5 Flash for TAPIDOR.\n",
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

        # Try API_BASE first, fallback to API_PROXY
        for endpoint in [f"{API_BASE}/ads/all", API_PROXY]:
            try:
                r = self.s.get(endpoint, params=params, timeout=30)
                if r.status_code == 200:
                    data = r.json()
                    return data.get("data", []) if isinstance(data, dict) else []
            except Exception as e:
                log.warning("Fetch listings from %s failed: %s", endpoint, e)

        return []

    def fetch_detail(self, slug: str) -> dict:
        url = f"{BASE_URL}/annonces/{slug}"
        r = self.s.get(url, headers={"Accept": "text/html"}, timeout=40, allow_redirects=True)
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
                result["date_parution"] = str(ad.get("date_parution", ""))[:10]
                result["meta_description"] = ad.get("meta_description", "")
                result["wilaya"] = (ad.get("wilaya") or {}).get("nom_fr", "")
                result["type_annonce"] = (ad.get("type_annonce") or {}).get("type_annonce_fr", "")
                result["sous_type"] = (ad.get("type_appel_offres") or {}).get("types_ao_fr", "")

                annonceurs = ad.get("annonceurs") or {}
                result["annonceur"] = annonceurs.get("annonceur") or ad.get("adresse_annonceur") or ""
                result["code_annonce"] = ad.get("code_annonce") or ""
                result["date_echeance"] = str(ad.get("date_echeance", ""))[:10] if ad.get("date_echeance") else ""
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

        m = re.search(r"Date d.{1,5}ch[eé]ance.*?(\d{4}-\d{2}-\d{2})", unescaped, re.DOTALL)
        if m and not result.get("date_echeance"):
            result["date_echeance"] = m.group(1)

        m = re.search(r'Entreprise concern[eé]e.*?</span>\s*<[^>]*>([^<]+)<', unescaped, re.DOTALL | re.IGNORECASE)
        if m and m.group(1).strip() and not result.get("entreprise_concernee"):
            result["entreprise_concernee"] = m.group(1).strip()

        # Montant extraction
        combined_text = f"{result.get('description', '')} {result.get('entreprise_concernee', '')}"
        m_montant = re.search(r'(\d[\d\s,.]*\s*(?:DA|DZD|dinars?))\b', combined_text, re.IGNORECASE)
        if m_montant:
            result["montant"] = m_montant.group(1).strip()

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
        date_val = str(ad_item.get("date_parution", ""))[:10]
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

        return urls

    def scrape_notice_type(self, notice_type: str, run_id: str, since_date: str = "") -> list[dict]:
        """Scrape either 'appels-doffres' or 'avis-attribution' with scan downloads and AI extraction."""
        label = "Appels d'Offres" if notice_type == "appels-doffres" else "Avis d'Attribution"
        log.info("--- Scraping %s (since date: %s) ---", label, since_date or "any")

        all_annonces = []
        max_pages_limit = 50 if since_date else MAX_PAGES

        for page in range(1, max_pages_limit + 1):
            listings = self.fetch_listings(page, notice_type=notice_type)
            if not listings:
                log.info("  Page %d: 0 listings -- stopping", page)
                break

            page_dates = [str(a.get("date_parution", ""))[:10] for a in listings if a.get("date_parution")]
            log.info("  Page %d: %d listings (dates: %s to %s)",
                     page, len(listings), max(page_dates) if page_dates else "?", min(page_dates) if page_dates else "?")

            if since_date:
                valid_on_page = [a for a in listings if str(a.get("date_parution", ""))[:10] >= since_date]
                all_annonces.extend(valid_on_page)
                if any(str(a.get("date_parution", ""))[:10] < since_date for a in listings):
                    log.info("  Reached ads older than cutoff %s at page %d -- stopping pagination", since_date, page)
                    break
            else:
                all_annonces.extend(listings)

            time.sleep(0.4)

        log.info("Total %s fetched (>= %s): %d", label, since_date or "all", len(all_annonces))

        # Filter by domain keywords with negative exclusions
        initial_matches = [a for a in all_annonces if KEYWORD_FILTER.search(a.get("titre", ""))]
        filtered = []
        for a in initial_matches:
            t = a.get("titre", "")
            if NON_GAZON_EXCLUSIONS.search(t) and not GAZON_EXPLICIT_KEYWORDS.search(t):
                log.info("  -> Pre-filter excluded non-gazon notice: %s", t[:70])
                continue
            filtered.append(a)
        log.info("After TAPIDOR gazon-exclusive filter: %d matching %s", len(filtered), label)

        today_str = datetime.now().strftime("%Y-%m-%d")
        results = []

        for a in filtered:
            ann_id = str(a.get("id_annonce", ""))
            slug = a.get("slug", "")
            titre = a.get("titre", "")
            raw_date = str(a.get("date_parution", ""))[:10] or today_str

            detail = {}
            for attempt in range(2):
                try:
                    detail = self.fetch_detail(slug) if slug else {}
                    if detail and not detail.get("_error"):
                        break
                except Exception as e:
                    time.sleep(2)

            # ── 1. Download Attached Scans & Upload to Google Drive ──
            scan_urls = self.get_scan_urls(a, detail.get("_html", ""))
            target_date = detail.get("date_parution") or raw_date
            ad_scan_dir = SCANS_DIR / target_date / ann_id
            ad_scan_dir.mkdir(parents=True, exist_ok=True)
            downloaded_scans = []

            for u in scan_urls:
                fname = u.split("/")[-1].split("?")[0]
                local_fp = ad_scan_dir / fname
                if not local_fp.exists():
                    try:
                        img_r = self.s.get(u, timeout=30)
                        if img_r.status_code == 200:
                            local_fp.write_bytes(img_r.content)
                            log.info("  -> Downloaded scan: %s (%d bytes)", fname, len(img_r.content))
                    except Exception as e:
                        log.warning("  -> Failed to download scan %s: %s", u, e)

                if local_fp.exists():
                    downloaded_scans.append(local_fp)
                    # Attempt Drive upload
                    if self.drive_mgr:
                        self.drive_mgr.upload_scan(local_fp, target_date)

            # ── 2. AI Scan Vision (NVIDIA NIM / Gemini OCR) ───────────
            ai_data = {}
            if downloaded_scans and (NVIDIA_API_KEY or GEMINI_API_KEY):
                engine_name = f"NVIDIA NIM ({NVIDIA_MODEL})" if NVIDIA_API_KEY else "Gemini"
                log.info("  -> Scanning attached newspaper file (%s) with %s...", downloaded_scans[0].name, engine_name)
                try:
                    ai_data = analyze_scan(downloaded_scans[0].read_bytes())
                    # If first scan did not contain budget/attribution and multiple pages exist, inspect second scan
                    if len(downloaded_scans) > 1 and (not ai_data.get("budget") or ai_data.get("budget") == "/"):
                        log.info("  -> Checking second scan page (%s) for complementary table data...", downloaded_scans[1].name)
                        ai_data_p2 = analyze_scan(downloaded_scans[1].read_bytes())
                        for k, v in ai_data_p2.items():
                            if v and v != "/" and (not ai_data.get(k) or ai_data.get(k) == "/"):
                                ai_data[k] = v
                except Exception as e:
                    log.warning("  -> AI Vision analysis error: %s", e)

            # ── 2.1 Gazon Scope Gate ────────────────────────────────
            # If AI Vision analyzed the scan and determined it's NOT a gazon project, discard it!
            if ai_data and ai_data.get("is_gazon") is False:
                log.info("  -> [REJECTED NON-GAZON] %s (AI confirmed no turf/gazon scope) -- Skipping.", titre[:65])
                continue

            # ── 3. Merge Web Data with AI Complementary Data ────────
            wilaya_val = (
                detail.get("wilaya")
                or (a.get("wilaya") or {}).get("nom_fr", "")
                or (ai_data.get("wilaya", "").strip().upper() if ai_data.get("wilaya") != "/" else "")
            )
            annonceur_val = (
                detail.get("annonceur")
                or (ai_data.get("annonceur", "").strip() if ai_data.get("annonceur") != "/" else "")
            )

            commune_val = parse_commune(titre, annonceur_val)
            if (not commune_val or commune_val == "/" or len(commune_val) <= 2) and ai_data.get("commune") and ai_data.get("commune") != "/":
                commune_val = ai_data.get("commune").strip().upper()

            entreprise_val = detail.get("entreprise_concernee", "")
            if (not entreprise_val or entreprise_val == "/") and ai_data.get("entreprise_attributaire") and ai_data.get("entreprise_attributaire") != "/":
                entreprise_val = ai_data.get("entreprise_attributaire").strip()

            montant_val = detail.get("montant", "")
            if (not montant_val or montant_val == "/") and ai_data.get("budget") and ai_data.get("budget") != "/":
                montant_val = ai_data.get("budget").strip()

            delai_val = parse_delai(detail.get("nbr_jours", 0), detail.get("description", ""))
            if (not delai_val or delai_val == "/") and ai_data.get("delai") and ai_data.get("delai") != "/":
                delai_val = ai_data.get("delai").strip()

            action_val = classify_action(titre)
            if action_val == "RÉALISATION" and ai_data.get("action") and ai_data.get("action") != "/":
                ai_act = classify_action(ai_data.get("action"))
                if ai_act != "RÉALISATION":
                    action_val = ai_act

            ptype_val = classify_type_projet(titre)
            if ptype_val == "COMPLEXE SPORTIF" and ai_data.get("type_projet") and ai_data.get("type_projet") != "/":
                ai_p = classify_type_projet(ai_data.get("type_projet"))
                if ai_p != "COMPLEXE SPORTIF":
                    ptype_val = ai_p

            row = {
                "run_id": run_id,
                "id": ann_id,
                "slug": slug,
                "url": f"{BASE_URL}/annonces/{slug}",
                "titre": titre,
                "action": action_val,
                "ptype": ptype_val,
                "wilaya": wilaya_val,
                "commune": commune_val,
                "date_parution": detail.get("date_parution") or raw_date,
                "date_parution_ao": ai_data.get("date_parution_ao", "/"),
                "date_echeance": detail.get("date_echeance", ""),
                "type_annonce": notice_type,
                "annonceur": annonceur_val,
                "code_annonce": detail.get("code_annonce", ""),
                "entreprise_concernee": entreprise_val,
                "montant": montant_val,
                "delai": delai_val,
                "description": detail.get("description", ""),
                "nbr_jours": detail.get("nbr_jours", 0),
                "scans": [str(p) for p in downloaded_scans],
                "scraped_at": datetime.now().isoformat(),
            }
            results.append(row)
            log.info("  [MATCH] %s | %s | Wilaya: %s | Commune: %s | Annonceur: %s",
                     ann_id, titre[:35], row["wilaya"], row["commune"], row["annonceur"][:25])
            time.sleep(0.5)

        return results

    # ── Main Run ─────────────────────────────────────────────────────────────

    def run(self, since_date: str | None = None):
        run_id = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        if not since_date:
            since_date = os.getenv("AM_SINCE_DATE", "").strip()
        if not since_date:
            # Default to last 7 days so weekend publications and downtime gaps are automatically recovered
            since_date = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d")

        log.info("=" * 60)
        log.info("AlgerieMarches Scraper Run %s started", run_id)
        log.info("Cutoff date (since): %s", since_date)
        log.info("Keyword filter: %s", ACTIVE_KEYWORDS)

        self.ensure_session()

        # 1. Scrape Appels d'Offres -> Target: APPEL D'OFFRE 2026
        log.info("\n>>> CYCLE 1: APPELS D'OFFRES <<<")
        ao_results = self.scrape_notice_type("appels-doffres", run_id, since_date=since_date)
        if ao_results:
            added_gsheet = append_to_gsheet(ao_results, notice_type="appels-doffres")
            log.info("AO Sync summary: %d added to Google Sheet", added_gsheet)

        # 2. Scrape Avis d'Attribution -> Target: AVIS D'ATTRIBUTIONS 2026
        log.info("\n>>> CYCLE 2: AVIS D'ATTRIBUTION <<<")
        aa_results = self.scrape_notice_type("avis-attribution", run_id, since_date=since_date)
        if aa_results:
            added_gsheet = append_to_gsheet(aa_results, notice_type="avis-attribution")
            log.info("AA Sync summary: %d added to Google Sheet", added_gsheet)

        # 3. Generate daily scans gallery for GitHub
        all_matches = (ao_results or []) + (aa_results or [])
        if all_matches:
            generate_daily_scans_summary(all_matches)

        self._save_cookies()
        log.info("\nRun %s successfully completed!", run_id)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="AlgerieMarches Scraper & Google Sheets Sync")
    parser.add_argument("--since", dest="since_date", default=None,
                        help="Cutoff publication date YYYY-MM-DD (defaults to yesterday)")
    args = parser.parse_args()

    scraper = AlgerieMarchesScraper()
    scraper.run(since_date=args.since_date)
