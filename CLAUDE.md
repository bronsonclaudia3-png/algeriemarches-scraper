# CLAUDE.md — Agent & Developer Guide for AlgerieMarches Scraper (TAPIDOR)

## 📌 Project Overview
Automated dual-pipeline scraper and AI vision OCR analyzer for Algerian Public Procurement (*Marchés Publics* on [AlgerieMarches.com](https://algeriemarches.com)). Tailored exclusively for **TAPIDOR**, an Algerian enterprise working on turf (*gazon synthétique*, *pelouse*, *engazonnement*) and on carpets / moquette for mosques and buildings.

---

## 🚨 Critical Business Rules (Non-Negotiable)

### 1. Strict Scope: GAZON and TAPIS only
TAPIDOR works on exactly two product families. A notice is kept only when its text or attached scan **EXPLICITLY** names one of them (`detect_scope()` returns `"GAZON"` or `"TAPIS"`; the AI returns `"scope"` plus a verbatim `"scope_evidence"` that must itself pass `detect_scope()`).

* **GAZON** — turf, synthetic or natural:
  * `GAZON`, `PELOUSE`, `ENGAZONNEMENT` / `ENGAZONNER`, `REVÊTEMENT EN GAZON` / `POSE DE GAZON` (plurals and `gazonné(e)(s)` included)
  * `عشب` / `تعشيب` / `نجيل` / `تكسية اصطناعية` — matched with attached prefixes (`العشب`, `بالعشب`, `التعشيب`)
  * Not turf projects: `tondeuse à gazon`, `coupe-gazon`, `débroussailleuse`, `tonte du gazon`, matériel agricole.
* **TAPIS** — carpets and moquette for **mosques and buildings** (offices, schools, hotels, conference rooms):
  * `TAPIS`, `MOQUETTE`, `MOQUETTAGE`, `زرابي` / `زربية`, `سجاد`, `موكيت`
  * Not carpets: `tapis roulant`, `tapis de course`, convoyeurs, `tapis d'enrobé` / `tapis en béton bitumineux` / any tapis next to enrobé/bitume/chaussée (road wearing course), sports mats (`tatami`, `tapis de judo/gym/lutte`, `sol sportif`), `tapis de souris`. Sports mats were explicitly ruled out by TAPIDOR (Oct 2026).
  * `tapis en gazon synthétique` is GAZON.
* **Never assume scope** from `stade`, `terrain de sport`, `ملعب`, `aire de jeux`, `mosquée`, `peinture`, `éclairage`, or `gradins`. Such notices are only sent to OCR (`prefilter_decision()` → `"ocr"`); civil masonry, school rehabilitation, roadworks and asphalt/MATICO grounds without turf **MUST be excluded**.
* Explicit text wins: if the title/description names gazon or tapis, an AI `"NONE"` does not reject the notice.
* If a scan cannot be read (every OCR engine failed), the notice is logged `[UNVERIFIED]` and retried on the next run — never rejected for good.

### 2. Publication Date Integrity (`date_parution_ao`)
* In **`AVIS D'ATTRIBUTIONS 2026`**, the `DATE DE PARUTION` column (Col F) records the date when the original tender (*Appel d'Offres*) was published in national newspapers.
* **Extraction Rule**: Extract `date_parution_ao` **ONLY** if explicitly written inside the text paragraph body (e.g., *"paru dans les journaux le 08/06/2026"* or *"الصادرة بتاريخ 2026/06/08 في جريدة..."*).
* **Strict Warning**: **NEVER invent a date** and **NEVER use the newspaper printing date or footer date** at the bottom (e.g., An-Nasr date, ANEP footer). If not mentioned in the body paragraph (as with direct consultations / *استشارة*), it **MUST be `"/"`**.
* `clean_date_parution_ao()` also drops dates that are not real dates, are after the award date, or are more than two years before it.

### 3. AI Model Restrictions
* **Permitted Vision Models (via NVIDIA NIM)**:
  * Primary: `z-ai/glm-5.3-flash` (60 s timeout; skipped for the rest of a run after 3 consecutive failures)
  * `moonshotai/kimi-k3` was retired by NVIDIA (HTTP 410, Oct 2026) and removed.
* **Permitted Google Gemini Fallback**: `gemini-2.5-flash`, then `gemini-flash-latest` (approved by TAPIDOR, Oct 2026). 429/503/timeouts are retried with backoff. `gemini-2.0-flash` / `gemini-1.5-flash` return 404 and were removed.
* **STRICTLY PROHIBITED**:
  * **DO NOT USE LLAMA**: Llama models hallucinate non-existent dates, output invalid JSON, and produce low OCR accuracy on bilingual French/Arabic scans.
  * **DO NOT USE OPUS**: Extremely high credit consumption; drains API balance within minutes.

---

## 📊 Google Sheets Contract & Schemas

**Spreadsheet ID**: `1y5-OxNeL_zBhCUNNEVoh1nKZVsyBYrq5EJcgKBEt920`

### 1. Worksheet: `APPEL D'OFFRE 2026`
* **Column Offset**: **2 leading empty columns** (Cols A & B are empty). Data starts at **Column C**.
* **Columns (C through N)**:
  1. `C` (Col 1): `N°` — Sequential counter integer (auto-increments from previous row).
  2. `D` (Col 2): `DATE` — Publication date in **`DD/MM/YYYY`** format.
  3. `E` (Col 3): `TITRE D'APPEL D'OFFRE` — Classified action verb (e.g., `RÉALISATION`, `AMÉNAGEMENT ET REVÊTEMENT`).
  4. `F` (Col 4): `Nombre de projet` — Static integer `1`. (Formatted as number `#`, never date).
  5. `G` (Col 5): `TYPE DE PROJET` — Facility type (e.g., `STADE DE FOOTBALL`, `STADE DE PROXIMITÉ`, `GAZON SYNTHÉTIQUE`, `TAPIS MOSQUÉE`, `MOQUETTE`, `TAPIS`).
  6. `H` (Col 6): `DATE DE PARUTION` — Tender publication date in **`DD/MM/YYYY`**.
  7. `I` (Col 7): `DATE D'ECHEANCE` — Tender deadline in **`DD/MM/YYYY`**.
  8. `J` (Col 8): `ANNONCEUR` — Contracting public authority (e.g., `COMMUNE`, `DJS DE LA WILAYA`, `DEP DE LA WILAYA`).
  9. `K` (Col 9): `WILAYA` — Wilaya name in UPPERCASE (e.g., `ORAN`, `SÉTIF`).
  10. `L` (Col 10): `COMMUNE` — Commune name in UPPERCASE (e.g., `ES SENIA`, `SALAH BEY`).
  11. `M` (Col 11): `CATEGORIE` — Static string `TRAVAUX PUBLICS`.
  12. `N` (Col 12): `SCANS` — Formula hyperlink: `=HYPERLINK("https://github.com/bronsonclaudia3-png/algeriemarches-scraper/tree/main/data/scans/YYYY-MM-DD/ID", "📁 Voir X Scans")`. Notices without a scan get `=HYPERLINK("https://algeriemarches.com/annonces/<slug>#ad-ID", "🔗 Annonce")` so the ID is still tracked.
* One row per notice.

> [!IMPORTANT]
> **Dates are written as forced text.** The spreadsheet locale is **`en_US`**, so a `USER_ENTERED` `02/11/2026` is read as *February 11* and shows up as a serial (`46064`). Every date cell goes through `sheet_text()` (leading apostrophe) — AO dates as `DD/MM/YYYY` via `format_date_dmy()`, AA dates in the formats below.

> [!IMPORTANT]
> **Rows are written with `ws.update()` to an explicit range starting at the `N°` column** (`C` for AO, `A` for AA), never with `append_rows()`. `values.append` detected the AO table at column C and shifted the blank-padded row two columns right (row 259, N° 277, Oct 2026). `read_sheet_state()` refuses to write (returns `-1`) if row 3 has no `N°` header or the `SCANS` header is not where expected. The scraper no longer writes any header cell (it used to overwrite `L3` "COMMUNE" with "SCANS").

### 2. Worksheet: `AVIS D'ATTRIBUTIONS 2026`
* **Column Offset**: **0 leading columns**. Data starts directly at **Column A**.
* **Columns (A through P)**:
  1. `A`: `N°` — Sequential counter integer.
  2. `B`: `DATE` — Ingestion date (`YYYY-MM-DD`).
  3. `C`: `TITRE D'AVIS D'ATTRIBUTION` — Action verb.
  4. `D`: `Nombre de projet` — Static integer `1`.
  5. `E`: `TYPE DE PROJET` — Facility type.
  6. `F`: `DATE DE PARUTION` — Original AO publication date from scan paragraph (`DD/MM/YYYY`) or `"/"`.
  7. `G`: `DATE D'ECHEANCE` — Legal contestation deadline date.
  8. `H`: `ANNONCEUR` — Public authority.
  9. `I`: `WILAYA` — Wilaya name.
  10. `J`: `COMMUNE` — Commune name.
  11. `K`: `DATE D'ATTRIBUTION` — Award decision date (`YYYY-MM-DD`).
  12. `L`: `BUDGET` — Corrected winning amount TTC (e.g., `19,736,328.50`).
  13. `M`: `ATTRIBUTION PAR` — Winning enterprise name (e.g., `CHELLALI Abdelghani`).
  14. `N`: `DELAI` — Execution delay (e.g., `03 MOIS`, `45 JOURS`).
  15. `O`: `WILAYA 2` — Static `"/"`.
  16. `P`: `SCANS` — Formula hyperlink to GitHub scan folder (or `#ad-ID` annonce link, as for AO).
* **One row per in-scope lot.** A multi-lot award gets one row per lot with that lot's own budget, winner and délai (`build_aa_rows()`). Lots are never summed; a lot without an amount gets `"/"`, never the notice total. Out-of-scope lots (e.g. masonry lots of a school project) are not written.
* `B`, `G` and `K` stay `YYYY-MM-DD`, `F` stays `DD/MM/YYYY` — all as forced text.

---

## 🔄 Deduplication Rules

The deduplication engine in `append_to_gsheet()` works **per notice** (all lots of one award are kept or skipped together):

1. **Primary: Ad Notice ID (`ann_id`)**:
   * Scans all cells and formulas in the Google Sheet for 5–6 digit IDs: `/(\d{5,6})(?:/|"|\.|$)` (GitHub scan path) and `#ad-(\d{5,6})` (annonce link). See `extract_ad_ids()`.
   * Any scraped item matching an ID in the sheet is **immediately skipped**.
2. **Secondary / Fallback (Appels d'Offres)**:
   * Checked against normalized publication date or deadline:
     * `("AO_PARUTION", action, ptype, norm_w, norm_c, item_parution)`
     * `("AO_ECHEANCE", norm_w, norm_c, item_echeance)`
   * **WARNING**: Never use a date-less `("NORM_LOC", norm_w, norm_c)` deduplicator. Tenders from the same commune on different dates in the same year are separate legitimate projects.
3. **Secondary / Fallback (Avis d'Attribution)**, per lot row:
   * `("AA_BUDGET", norm_w, norm_c, "<amount to 2 decimals>")` — the same exact amount in the same commune (catches the French and Arabic versions of one award).
   * `("AA_NOBUDGET", norm_w, norm_c, norm_winner, award_date)` when there is no amount.

### Verdict cache (`state/verdicts.json`)
* Every notice that reaches OCR gets a verdict (`match` with its result row, or `reject` with a reason), so each notice is OCR'd **once** instead of on every daily run of its 7-day window. Committed by GitHub Actions; entries older than 60 days are pruned.
* Unverified notices (OCR down) are **not** cached.
* **Bump `VERDICTS_VERSION`** whenever the scope rules or the OCR prompt change, so old verdicts are re-checked.

---

## 🛠️ Developer Commands

### Environment Setup
```bash
# Windows
.venv\Scripts\activate
# Linux/macOS
source .venv/bin/activate

pip install -r requirements.txt
```

### Running Tests
```bash
# Run all unit tests
python -m unittest discover tests -v

# Run individual test modules
python -m unittest tests/test_classifiers.py
python -m unittest tests/test_dedup_and_dates.py
python -m unittest tests/test_hardening.py   # scope, amounts, Sheet writer (fake worksheet), OCR chain, crawl
```

### Running the Scraper
```bash
# Full daily run (both AO and AA)
python scraper.py

# Run specific notice type
python scraper.py --type appels-doffres
python scraper.py --type avis-attribution

# Run with date catch-up
python scraper.py --since 2026-10-01

# Syntax verification
python -m py_compile scraper.py
```

`scraper.py` exits with code **1** (red GitHub Actions run) when a listing page fails after retries, the page cap is hit before the cutoff date, a notice raises an exception, or the Sheet write fails. The commit step still runs (`if: always()`).

### Testing Google Sheet Sync Connection (Dry Run)
```python
python -c "from scraper import append_to_gsheet; append_to_gsheet([], 'appels-doffres'); append_to_gsheet([], 'avis-attribution')"
```

---

## 🗂️ Key Files & Architecture

| File / Folder | Role |
| :--- | :--- |
| `scraper.py` | Monolithic engine: auth, listing crawl, detail extraction, scan download, AI OCR & Google Sheets sync. |
| `tests/` | Unit test suite covering regex classifiers, date normalizers, and deduplication logic. |
| `data/scans/` | Local and GitHub-hosted archive of downloaded newspaper scans organized by `YYYY-MM-DD/<id>/`. |
| `state/cookies.json` | Active AlgerieMarches session cookies. Auto-refreshed and committed by GitHub Actions. |
| `state/verdicts.json` | Per-notice OCR verdict cache (see Deduplication Rules). Committed by GitHub Actions. |
| `state/service_account.json` | Google Cloud Service Account key for Google Sheets API (git-ignored). |
| `.github/workflows/scraper.yml` | GitHub Actions workflow running daily at **07:00 UTC (08:00 Algerian Time)**. |
| `.env` | Local environment variables (`AM_EMAIL`, `AM_PASSWORD`, `NVIDIA_API_KEY`, `GEMINI_API_KEY`, etc.). |

---

## ⚠️ Known Gotchas & Workarounds

1. **Windows PowerShell Console Encoding**:
   * Printing accented French characters on Windows PowerShell can raise `UnicodeEncodeError: 'charmap' codec can't encode character`.
   * **Fix**: In standalone scripts, always add at the top:
     ```python
     import sys
     sys.stdout.reconfigure(encoding='utf-8')
     ```
2. **Session Keep-Alive**:
   * AlgerieMarches expires sessions after inactivity. `scraper.py` handles session check, CSRF retrieval, and auto-login with session promotion.
   * `state/cookies.json` is committed back to GitHub after every run to keep the remote runner authenticated.
3. **Google Sheets Formula Escapes**:
   * Hyperlinks are injected as `=HYPERLINK("...", "...")`. When writing raw ranges via `gspread`, ensure `value_input_option="USER_ENTERED"` is specified so formulas are evaluated by Google Sheets.
4. **API timestamps are UTC**: `date_parution` comes as `2026-10-22T00:00:00.000Z`. Use `to_local_date()` (Algeria, UTC+1), never `str(x)[:10]`.
5. **Listing volume**: ~950-1,000 notices per type per 7 days. `MAX_PAGES_CATCHUP` (120 pages) must stay well above that; the old 50-page cap was hit on 2026-10-22.
6. **Broken "à" in titles**: AlgerieMarches titles contain `U+FFFD U+00A0` instead of `à`; `normalize_text()` repairs it before matching.
