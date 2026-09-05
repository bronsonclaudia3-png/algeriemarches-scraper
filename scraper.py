"""
AlgerieMarches Scraper v4 — Dual Scrape (Appels d'Offres & Avis d'Attributions)
Automated daily sync to Google Sheets & local Downloads Excel file.

Targets:
  1. "APPEL D'OFFRE 2026" (11 columns)
  2. "AVIS D'ATTRIBUTIONS 2026" (15 columns)
Domain:
  Artificial turf, stadiums, sports facilities, playgrounds, school courtyards (TAPIDOR).
"""

import csv
import json
import logging
import os
import re
import sys
import time
import unicodedata
from copy import copy
from datetime import datetime
from pathlib import Path
from urllib.parse import unquote

import gspread
import openpyxl
import requests
from dotenv import load_dotenv

# ── Paths ────────────────────────────────────────────────────────────────────
SCRIPT_DIR = Path(__file__).resolve().parent
STATE_DIR = SCRIPT_DIR / "state"
DATA_DIR = SCRIPT_DIR / "data"
COOKIE_FILE = STATE_DIR / "cookies.json"
LOG_FILE = SCRIPT_DIR / "scraper.log"

STATE_DIR.mkdir(exist_ok=True)
DATA_DIR.mkdir(exist_ok=True)

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

BASE_URL = "https://algeriemarches.com"
API_BASE = "https://api.algeriemarches.com/api"
API_PROXY = f"{BASE_URL}/api/proxy/ads"

CSRF_URL = f"{BASE_URL}/api/auth/csrf"
LOGIN_URL = f"{BASE_URL}/api/auth/callback/credentials"
SESSIONS_URL = f"{BASE_URL}/api/auth/active-sessions"
DISCONNECT_URL = f"{BASE_URL}/api/proxy/auth/disconnect-session"

PAGE_SIZE = 20
MAX_PAGES = 3

# Sport & turf keywords for TAPIDOR
DEFAULT_KEYWORDS = r"\b(?:gazon|pelouse|engazonnement|stade|terrain|sport|football|jeux|matico|matiquo|athletisme|cour|cours)\b"
KEYWORDS_ENV = os.getenv("AM_KEYWORDS", "").strip()
ACTIVE_KEYWORDS = KEYWORDS_ENV if KEYWORDS_ENV else DEFAULT_KEYWORDS
KEYWORD_FILTER = re.compile(ACTIVE_KEYWORDS, re.IGNORECASE)

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

def parse_budget(montant_str: str):
    if not montant_str:
        return "/"
    clean = re.sub(r"[^\d.,]", "", str(montant_str)).replace(",", ".")
    try:
        val = float(clean)
        return int(val) if val.is_integer() else val
    except Exception:
        return "/"

def parse_delai(nbr_jours: int, description: str) -> str:
    if nbr_jours and int(nbr_jours) > 0:
        return f"{nbr_jours} JOURS"
    m = re.search(r"\b(\d+\s*(?:JOURS?|MOIS))\b", str(description).upper())
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

def get_latest_downloads_excel() -> Path | None:
    downloads = Path.home() / "Downloads"
    if not downloads.exists():
        return None
    files = [f for f in downloads.glob("*.xlsx") if not f.name.startswith("~$") and not f.name.endswith("_BACKUP.xlsx")]
    if not files:
        return None
    files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    return files[0]

# ── Excel Local Sync ─────────────────────────────────────────────────────────

def append_to_excel(results: list[dict], notice_type: str, excel_path: Path | None = None) -> int:
    """
    Append results to the matching sheet in the Downloads Excel file.
    - notice_type 'appels-doffres' -> 'APPEL D'OFFRE 2026' (11 columns)
    - notice_type 'avis-attribution' -> 'AVIS D'ATTRIBUTIONS 2026' (15 columns)
    """
    if excel_path is None:
        excel_path = get_latest_downloads_excel()

    if not excel_path or not excel_path.exists():
        log.warning("No Excel file found in Downloads to update.")
        return 0

    target_sheet = "APPEL D'OFFRE 2026" if notice_type == "appels-doffres" else "AVIS D'ATTRIBUTIONS 2026"
    log.info("Checking Downloads Excel file: %s (Target: %s)", excel_path.name, target_sheet)

    try:
        wb = openpyxl.load_workbook(excel_path)
    except PermissionError:
        log.warning("File %s is currently open. Saving to updated copy.", excel_path.name)
        excel_path = excel_path.parent / f"{excel_path.stem}_updated.xlsx"
        if excel_path.exists():
            wb = openpyxl.load_workbook(excel_path)
        else:
            return 0
    except Exception as e:
        log.error("Failed to open workbook %s: %s", excel_path, e)
        return 0

    if target_sheet not in wb.sheetnames:
        log.warning("Sheet '%s' not found in %s", target_sheet, excel_path.name)
        return 0

    ws = wb[target_sheet]

    # Read existing rows to deduplicate and find next N°
    last_row_idx = 3
    last_num = 0
    existing_items = set()

    for r in range(4, 1500):
        val_num = ws.cell(row=r, column=1).value
        if val_num is not None:
            last_row_idx = r
            try:
                last_num = int(val_num)
            except Exception:
                pass
            if notice_type == "appels-doffres":
                action = str(ws.cell(row=r, column=3).value or "").strip().upper()
                t_proj = str(ws.cell(row=r, column=5).value or "").strip().upper()
                wilaya = str(ws.cell(row=r, column=9).value or "").strip().upper()
                commune = str(ws.cell(row=r, column=10).value or "").strip().upper()
                existing_items.add((action, t_proj, wilaya, commune))
            else:
                attr_par = str(ws.cell(row=r, column=13).value or "").strip().upper()
                wilaya = str(ws.cell(row=r, column=9).value or "").strip().upper()
                t_proj = str(ws.cell(row=r, column=5).value or "").strip().upper()
                existing_items.add((attr_par, wilaya, t_proj))
        else:
            # Verify if trailing rows exist
            empty = True
            for check in range(r, min(r + 10, 1500)):
                if ws.cell(row=check, column=1).value is not None:
                    empty = False
                    break
            if empty:
                break

    log.info("Sheet '%s' has %d rows (last N°: %d)", target_sheet, last_num, last_num)
    ref_row_idx = last_row_idx if last_row_idx >= 4 else 4

    added_count = 0
    today_dt = datetime(datetime.now().year, datetime.now().month, datetime.now().day)

    for item in results:
        titre = item.get("titre", "")
        annonceur = item.get("annonceur", "")
        action = classify_action(titre)
        ptype = classify_type_projet(titre)
        wilaya = (item.get("wilaya") or "").strip().upper()
        commune = parse_commune(titre, annonceur)
        ann_val = clean_annonceur(annonceur)

        dt_parution = parse_date(item.get("date_parution", ""))
        dt_echeance = parse_date(item.get("date_echeance", ""))

        if notice_type == "appels-doffres":
            dedup_key = (action.upper(), ptype.upper(), wilaya, commune.upper())
            if dedup_key in existing_items:
                log.info("  Already in Excel (AO): %s | %s (%s)", action, ptype, wilaya)
                continue

            last_num += 1
            last_row_idx += 1
            added_count += 1
            existing_items.add(dedup_key)

            row_values = [
                last_num,                                # Col 1: N°
                today_dt,                                # Col 2: DATE
                action,                                  # Col 3: TITRE D'APPEL D'OFFRE
                1,                                       # Col 4: Nombre de projet
                ptype,                                   # Col 5: TYPE DE PROJET
                dt_parution,                             # Col 6: DATE DE PARUTION
                dt_echeance,                             # Col 7: DATE D'ECHEANCE
                ann_val,                                 # Col 8: ANNONCEUR
                wilaya if wilaya else "/",               # Col 9: WILAYA
                commune,                                 # Col 10: COMMUNE
                "TRAVAUX PUBLICS",                       # Col 11: CATEGORIE
            ]
        else:
            attr_par = item.get("entreprise_concernee", "").strip()
            dedup_key = (attr_par.upper(), wilaya, ptype.upper())
            if dedup_key in existing_items and attr_par != "":
                log.info("  Already in Excel (AA): %s (%s)", attr_par[:30], wilaya)
                continue

            last_num += 1
            last_row_idx += 1
            added_count += 1
            existing_items.add(dedup_key)

            budget = parse_budget(item.get("montant", ""))
            delai = parse_delai(item.get("nbr_jours", 0), item.get("description", ""))

            row_values = [
                last_num,                                # Col 1: N°
                today_dt,                                # Col 2: DATE
                action,                                  # Col 3: TITRE D'AVIS D'ATTRIBUTION
                1,                                       # Col 4: Nombre de projet
                ptype,                                   # Col 5: TYPE DE PROJET
                dt_parution,                             # Col 6: DATE DE PARUTION
                dt_echeance,                             # Col 7: DATE D'ECHEANCE
                ann_val,                                 # Col 8: ANNONCEUR
                wilaya if wilaya else "/",               # Col 9: WILAYA
                commune,                                 # Col 10: COMMUNE
                dt_parution,                             # Col 11: DATE D'ATTRIBUTION
                budget,                                  # Col 12: BUDGET
                attr_par if attr_par else "/",           # Col 13: ATTRIBUTION PAR
                delai,                                   # Col 14: DELAI
                "/",                                     # Col 15: WILAYA 2
            ]

        for col_idx, val in enumerate(row_values, start=1):
            cell = ws.cell(row=last_row_idx, column=col_idx, value=val)
            ref_cell = ws.cell(row=ref_row_idx, column=col_idx)
            if ref_cell.font: cell.font = copy(ref_cell.font)
            if ref_cell.border: cell.border = copy(ref_cell.border)
            if ref_cell.alignment: cell.alignment = copy(ref_cell.alignment)
            if ref_cell.number_format: cell.number_format = ref_cell.number_format

    if added_count > 0:
        ws.cell(row=1, column=4, value="=TODAY()")
        try:
            wb.save(excel_path)
            log.info("Successfully appended %d new rows to %s (sheet: %s)", added_count, excel_path.name, target_sheet)
        except PermissionError:
            alt_path = excel_path.parent / f"{excel_path.stem}_updated.xlsx"
            wb.save(alt_path)
            log.warning("File locked. Saved %d rows to: %s", added_count, alt_path.name)
    else:
        log.info("No new rows needed for Excel '%s' (all up to date)", target_sheet)

    return added_count

# ── Google Sheets Sync ───────────────────────────────────────────────────────

def append_to_gsheet(results: list[dict], notice_type: str, sheet_id: str | None = None) -> int:
    """
    Append results to the matching worksheet in Google Sheet.
    - notice_type 'appels-doffres' -> 'APPEL D'OFFRE 2026' (11 columns)
    - notice_type 'avis-attribution' -> 'AVIS D'ATTRIBUTIONS 2026' (15 columns)
    """
    sheet_id = sheet_id or os.getenv("GOOGLE_SHEET_ID", "1y5-OxNeL_z8hCUNNEVoh1nKZVsyBYrq5EJcgKBEt920")
    if not sheet_id:
        log.info("No GOOGLE_SHEET_ID configured -- skipping Google Sheet sync.")
        return 0

    sa_json_env = os.getenv("GCP_SERVICE_ACCOUNT_KEY")
    sa_file = SCRIPT_DIR / "state" / "service_account.json"

    gc = None
    if sa_json_env:
        try:
            creds_dict = json.loads(sa_json_env)
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
                existing_items.add((attr_par, wilaya, t_proj))

    log.info("Google Sheet '%s' has %d rows (last N°: %d)", target_ws.title, len(all_values) - 3, last_num)

    rows_to_append = []
    today_str = datetime.now().strftime("%Y-%m-%d")

    for item in results:
        titre = item.get("titre", "")
        annonceur = item.get("annonceur", "")
        action = classify_action(titre)
        ptype = classify_type_projet(titre)
        wilaya = (item.get("wilaya") or "").strip().upper()
        commune = parse_commune(titre, annonceur)
        ann_val = clean_annonceur(annonceur)

        dt_parution = str(item.get("date_parution", ""))[:10] or "/"
        dt_echeance = str(item.get("date_echeance", ""))[:10] or "/"

        if notice_type == "appels-doffres":
            dedup_key = (action.upper(), ptype.upper(), wilaya, commune.upper())
            if dedup_key in existing_items:
                log.info("  Already in Google Sheet (AO): %s | %s (%s)", action, ptype, wilaya)
                continue

            last_num += 1
            existing_items.add(dedup_key)

            row_data = [
                last_num,                                # Col 1: N°
                today_str,                               # Col 2: DATE
                action,                                  # Col 3: TITRE D'APPEL D'OFFRE
                1,                                       # Col 4: Nombre de projet
                ptype,                                   # Col 5: TYPE DE PROJET
                dt_parution,                             # Col 6: DATE DE PARUTION
                dt_echeance,                             # Col 7: DATE D'ECHEANCE
                ann_val,                                 # Col 8: ANNONCEUR
                wilaya if wilaya else "/",               # Col 9: WILAYA
                commune,                                 # Col 10: COMMUNE
                "TRAVAUX PUBLICS",                       # Col 11: CATEGORIE
            ]
            rows_to_append.append(row_data)
        else:
            attr_par = item.get("entreprise_concernee", "").strip()
            dedup_key = (attr_par.upper(), wilaya, ptype.upper())
            if dedup_key in existing_items and attr_par != "":
                log.info("  Already in Google Sheet (AA): %s (%s)", attr_par[:30], wilaya)
                continue

            last_num += 1
            existing_items.add(dedup_key)

            budget = parse_budget(item.get("montant", ""))
            delai = parse_delai(item.get("nbr_jours", 0), item.get("description", ""))

            row_data = [
                last_num,                                # Col 1: N°
                today_str,                               # Col 2: DATE
                action,                                  # Col 3: TITRE D'AVIS D'ATTRIBUTION
                1,                                       # Col 4: Nombre de projet
                ptype,                                   # Col 5: TYPE DE PROJET
                dt_parution,                             # Col 6: DATE DE PARUTION
                dt_echeance,                             # Col 7: DATE D'ECHEANCE
                ann_val,                                 # Col 8: ANNONCEUR
                wilaya if wilaya else "/",               # Col 9: WILAYA
                commune,                                 # Col 10: COMMUNE
                dt_parution,                             # Col 11: DATE D'ATTRIBUTION
                budget,                                  # Col 12: BUDGET
                attr_par if attr_par else "/",           # Col 13: ATTRIBUTION PAR
                delai,                                   # Col 14: DELAI
                "/",                                     # Col 15: WILAYA 2
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
        r = self.s.post(
            LOGIN_URL,
            json={
                "csrfToken": csrf,
                "identifier": EMAIL,
                "email": EMAIL,
                "password": PASSWORD,
                "callbackUrl": BASE_URL,
            },
            headers={"Content-Type": "application/json"},
            allow_redirects=False,
            timeout=40,
        )
        has_auth = self._has_auth_cookies()
        loc = r.headers.get("Location", "")
        log.info("Login: status=%d location=%s has_auth=%s", r.status_code, loc[:80], has_auth)
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
            raise RuntimeError("Login failed -- check password in .env")

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

        result = {"_source": "html", "_html_len": len(html)}

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

    def scrape_notice_type(self, notice_type: str, run_id: str) -> list[dict]:
        """Scrape either 'appels-doffres' or 'avis-attribution'."""
        label = "Appels d'Offres" if notice_type == "appels-doffres" else "Avis d'Attribution"
        log.info("--- Scraping %s (pages 1 to %d) ---", label, MAX_PAGES)

        all_annonces = []
        for page in range(1, MAX_PAGES + 1):
            listings = self.fetch_listings(page, notice_type=notice_type)
            if not listings:
                log.info("  Page %d: 0 listings -- stopping", page)
                break
            log.info("  Page %d: %d listings", page, len(listings))
            all_annonces.extend(listings)
            time.sleep(0.5)

        log.info("Total %s fetched: %d", label, len(all_annonces))

        # Filter by domain keywords (turf/sport/playgrounds)
        filtered = [a for a in all_annonces if KEYWORD_FILTER.search(a.get("titre", ""))]
        log.info("After TAPIDOR keyword filter: %d matching %s", len(filtered), label)

        results = []
        for a in filtered:
            ann_id = str(a.get("id_annonce", ""))
            slug = a.get("slug", "")
            titre = a.get("titre", "")
            raw_date = str(a.get("date_parution", ""))[:10]

            detail = {}
            for attempt in range(2):
                try:
                    detail = self.fetch_detail(slug) if slug else {}
                    if detail and not detail.get("_error"):
                        break
                except Exception as e:
                    time.sleep(2)

            row = {
                "run_id": run_id,
                "id": ann_id,
                "slug": slug,
                "url": f"{BASE_URL}/annonces/{slug}",
                "titre": titre,
                "wilaya": detail.get("wilaya") or (a.get("wilaya") or {}).get("nom_fr", ""),
                "date_parution": detail.get("date_parution") or raw_date,
                "date_echeance": detail.get("date_echeance", ""),
                "type_annonce": notice_type,
                "annonceur": detail.get("annonceur", ""),
                "code_annonce": detail.get("code_annonce", ""),
                "entreprise_concernee": detail.get("entreprise_concernee", ""),
                "montant": detail.get("montant", ""),
                "description": detail.get("description", ""),
                "nbr_jours": detail.get("nbr_jours", 0),
                "scraped_at": datetime.now().isoformat(),
            }
            results.append(row)
            log.info("  [MATCH] %s | %s | %s | %s", ann_id, titre[:35], row["wilaya"], row["annonceur"][:25])
            time.sleep(0.5)

        return results

    # ── Main Run ─────────────────────────────────────────────────────────────

    def run(self):
        run_id = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        log.info("=" * 60)
        log.info("AlgerieMarches Scraper Run %s started", run_id)
        log.info("Keyword filter: %s", ACTIVE_KEYWORDS)

        self.ensure_session()

        # 1. Scrape Appels d'Offres -> Target: APPEL D'OFFRE 2026
        log.info("\n>>> CYCLE 1: APPELS D'OFFRES <<<")
        ao_results = self.scrape_notice_type("appels-doffres", run_id)
        if ao_results:
            added_excel = append_to_excel(ao_results, notice_type="appels-doffres")
            added_gsheet = append_to_gsheet(ao_results, notice_type="appels-doffres")
            log.info("AO Sync summary: %d added to Excel, %d added to Google Sheet", added_excel, added_gsheet)

        # 2. Scrape Avis d'Attribution -> Target: AVIS D'ATTRIBUTIONS 2026
        log.info("\n>>> CYCLE 2: AVIS D'ATTRIBUTION <<<")
        aa_results = self.scrape_notice_type("avis-attribution", run_id)
        if aa_results:
            added_excel = append_to_excel(aa_results, notice_type="avis-attribution")
            added_gsheet = append_to_gsheet(aa_results, notice_type="avis-attribution")
            log.info("AA Sync summary: %d added to Excel, %d added to Google Sheet", added_excel, added_gsheet)

        self._save_cookies()
        log.info("\nRun %s successfully completed!", run_id)


if __name__ == "__main__":
    scraper = AlgerieMarchesScraper()
    scraper.run()
