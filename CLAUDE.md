# CLAUDE.md — Agent & Developer Guide for AlgerieMarches Scraper (TAPIDOR)

## 📌 Project Overview
Automated dual-pipeline scraper and AI vision OCR analyzer for Algerian Public Procurement (*Marchés Publics* on [AlgerieMarches.com](https://algeriemarches.com)). Tailored exclusively for **TAPIDOR**, a leading Algerian enterprise specializing strictly in sports turf (*gazon synthétique*, *pelouse synthétique*, *engazonnement*).

---

## 🚨 Critical Business Rules (Non-Negotiable)

### 1. Strict Sports Turf (Gazon) Exclusivity
* **TAPIDOR ONLY works on sports turf / gazon projects.**
* **Never assume a project is gazon simply because it mentions `stade`, `terrain de sport`, `ملعب`, `aire de jeux`, `peinture`, `éclairage`, or `gradins`.**
* Many Algerian tenders are for civil masonry, school rehabilitation, roadworks, municipal paving, or asphalt (MATICO) multi-sports grounds **WITHOUT ANY TURF**.
* **Rule**: Only classify an ad or scan as `is_gazon: true` if the text or attached scan **EXPLICITLY** mentions:
  * `GAZON` (synthétique ou naturel)
  * `PELOUSE` (synthétique ou naturelle)
  * `ENGAZONNEMENT` / `ENGAZONNER`
  * `REVÊTEMENT EN GAZON` / `POSE DE GAZON`
  * `عشب اصطناعي` / `عشب طبيعي` / `تعشيب` / `تكسية بالعشب الاصطناعي`
* If it is general construction, fencing, lighting, roadworks, or general rehabilitation without turf: **it MUST be excluded (`is_gazon: false`)**.

### 2. Publication Date Integrity (`date_parution_ao`)
* In **`AVIS D'ATTRIBUTIONS 2026`**, the `DATE DE PARUTION` column (Col F) records the date when the original tender (*Appel d'Offres*) was published in national newspapers.
* **Extraction Rule**: Extract `date_parution_ao` **ONLY** if explicitly written inside the text paragraph body (e.g., *"paru dans les journaux le 08/06/2026"* or *"الصادرة بتاريخ 2026/06/08 في جريدة..."*).
* **Strict Warning**: **NEVER invent a date** and **NEVER use the newspaper printing date or footer date** at the bottom (e.g., An-Nasr date, ANEP footer). If not mentioned in the body paragraph (as with direct consultations / *استشارة*), it **MUST be `"/"`**.

### 3. AI Model Restrictions
* **Permitted Vision Models (via NVIDIA NIM)**:
  * Primary: `z-ai/glm-5.3-flash`
  * Fallback: `moonshotai/kimi-k3`
* **Permitted Google Gemini Fallback**: `gemini-2.5-flash`
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
  5. `G` (Col 5): `TYPE DE PROJET` — Facility type (e.g., `STADE DE FOOTBALL`, `STADE DE PROXIMITÉ`, `GAZON SYNTHÉTIQUE`).
  6. `H` (Col 6): `DATE DE PARUTION` — Tender publication date in **`DD/MM/YYYY`**.
  7. `I` (Col 7): `DATE D'ECHEANCE` — Tender deadline in **`DD/MM/YYYY`**.
  8. `J` (Col 8): `ANNONCEUR` — Contracting public authority (e.g., `COMMUNE`, `DJS DE LA WILAYA`, `DEP DE LA WILAYA`).
  9. `K` (Col 9): `WILAYA` — Wilaya name in UPPERCASE (e.g., `ORAN`, `SÉTIF`).
  10. `L` (Col 10): `COMMUNE` — Commune name in UPPERCASE (e.g., `ES SENIA`, `SALAH BEY`).
  11. `M` (Col 11): `CATEGORIE` — Static string `TRAVAUX PUBLICS`.
  12. `N` (Col 12): `SCANS` — Formula hyperlink: `=HYPERLINK("https://github.com/bronsonclaudia3-png/algeriemarches-scraper/tree/main/data/scans/YYYY-MM-DD/ID", "📁 Voir X Scans")`.

> [!IMPORTANT]
> **Date Formatting in AO**: Always pass dates through `format_date_dmy()` so they write as `DD/MM/YYYY`. Unformatted or ISO dates in French Google Sheets can be converted into raw Excel integer serials (e.g., `46302`).

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
  16. `P`: `SCANS` — Formula hyperlink to GitHub scan folder.

---

## 🔄 Deduplication Rules

The deduplication engine in `append_to_gsheet()` uses a two-tier strategy:

1. **Primary: Ad Notice ID (`ann_id`)**:
   * Scans all cells and formulas in the Google Sheet for 5–6 digit IDs: `re.findall(r'/(\d{5,6})(?:/|"|\.|$)', str(cell))`.
   * Any scraped item matching an ID in `existing_ad_ids` is **immediately skipped**.
2. **Secondary / Fallback (Appels d'Offres)**:
   * Checked against normalized publication date or deadline:
     * `("AO_PARUTION", action, ptype, norm_w, norm_c, item_parution)`
     * `("AO_ECHEANCE", norm_w, norm_c, item_echeance)`
   * **WARNING**: Never use a date-less `("NORM_LOC", norm_w, norm_c)` deduplicator. Tenders from the same commune on different dates in the same year are separate legitimate projects.

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
