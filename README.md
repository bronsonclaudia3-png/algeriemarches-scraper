# 🇩🇿 AlgerieMarches Scraper & Analyzer (TAPIDOR)

> **Automated Dual Scraper & AI Vision Analysis for Algerian Public Procurement (Marchés Publics)**  
> Tailored exclusively for **TAPIDOR** — Specializing in Sports Turf (*Gazon Synthétique*, *Pelouse*, *Engazonnement* & Sports Facilities).

[![Daily Scraper](https://github.com/bronsonclaudia3-png/algeriemarches-scraper/actions/workflows/scraper.yml/badge.svg)](https://github.com/bronsonclaudia3-png/algeriemarches-scraper/actions/workflows/scraper.yml)
[![Python 3.12](https://img.shields.io/badge/Python-3.12-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-Private-red.svg)]()

---

## 📌 Overview

This repository automates the daily extraction, OCR analysis, filtering, and reporting of public tenders and contract awards published on **AlgerieMarches.com**.

### Target Datasets
1. **`APPEL D'OFFRE 2026`**: National and local tenders for sports facility projects, synthetic turf installations, and stadium rehabilitation.
2. **`AVIS D'ATTRIBUTIONS 2026`**: Winning bids, winning contractors (*entreprises attributaires*), final amounts (TTC), delays, and original tender publication dates.

---

## 🚀 Key Features

* **Strict Sports Turf (Gazon) Filtering**:
  * Precision keyword matching for turf and sports grounds (*gazon*, *pelouse*, *engazonnement*, *عشب اصطناعي*, *تعشيب*, *stade*, *terrain de sport*).
  * Robust negative exclusion patterns preventing civil works false positives (e.g. *glissement de terrain*, airport runways, paving/enrobé, dams, general administrative buildings).

* **AI Scan Vision & OCR Analysis**:
  * Automatically downloads attached newspaper clippings (ANEP scans) in French and Arabic.
  * Analyzes scans using **NVIDIA NIM** (`google/diffusiongemma-26b-a4b-it`, `z-ai/glm-5.3-flash`, `moonshotai/kimi-k3`) with automatic fallback to **Google Gemini Vision** (`gemini-2.5-flash`).
  * Extracts winning contractor name, corrected TTC amount, execution delays, municipal commune, and original tender publication date.

* **Anti-Hallucination & Evidence Verification**:
  * Requires explicit verbatim proof from scan text (`gazon_mention`) before accepting tenders that only mention generic keywords.
  * Discards invalid/invented publication dates (e.g., historical decree dates or municipal consultation dates).

* **Airtight Duplicate Prevention**:
  * Tracks unique Algerian tender notice IDs (`ann_id`) directly across Google Sheets formulas and cell values.
  * Dynamic column offset detection ensuring perfect alignment across worksheets.

* **Daily Automated Cloud Execution**:
  * Runs automatically every morning at **07:00 UTC (08:00 AM Algerian Time)** via GitHub Actions.
  * Syncs live results to the shared Google Sheet and updates the local GitHub scans archive.

---

## 📁 Repository Structure

```text
├── .github/workflows/
│   └── scraper.yml             # GitHub Actions daily cron schedule (07:00 UTC)
├── data/
│   ├── scans/                  # Downloaded newspaper scans organized by date
│   │   ├── YYYY-MM-DD/         # Date folder containing notice scans and README galleries
│   │   └── README.md           # Master gallery index of all archived scans
├── state/
│   ├── cookies.json            # Persistent authenticated session cookies
│   └── service_account.json    # Google Cloud service account key (git-ignored)
├── scraper.py                  # Core engine: scraping, auth, AI OCR & Google Sheets sync
├── requirements.txt            # Python dependencies
├── .env.example                # Template for required environment variables
└── README.md                   # Project documentation
```

---

## ⚙️ Configuration & Environment Variables

Create a `.env` file in the root directory with the following keys:

```env
# AlgerieMarches Credentials
AM_EMAIL=direction@tapidor.com
AM_PASSWORD=your_password_here

# AI Vision API Keys
NVIDIA_API_KEY=nvapi-...
GEMINI_API_KEY=AIzaSy...

# Google Sheets Integration
GOOGLE_SHEET_ID=1y5-OxNeL_zBhCUNNEVoh1nKZVsyBYrq5EJcgKBEt920
GCP_SERVICE_ACCOUNT_KEY={"type": "service_account", ...} # Or use state/service_account.json

# Optional: Google Drive Scan Backup Folder
GOOGLE_DRIVE_FOLDER_ID=your_gdrive_folder_id
```

---

## 🛠️ Local Installation & Execution

### 1. Prerequisites
* Python 3.12+
* Virtual environment (`venv` or `uv`)

### 2. Install Dependencies
```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# Linux/macOS
source .venv/bin/activate

pip install -r requirements.txt
```

### 3. Run Scraper
```bash
# Run daily sync for both Appels d'Offres and Avis d'Attribution
python scraper.py

# Run for a specific start date (e.g., catch-up)
python scraper.py --since 2026-09-25

# Run only Appels d'Offres
python scraper.py --type appels-doffres

# Run only Avis d'Attribution
python scraper.py --type avis-attribution
```

---

## 📊 Google Sheets Columns Reference

### 1. `APPEL D'OFFRE 2026`
| Col | Header | Description |
| :---: | :--- | :--- |
| **C** | `N°` | Incrementing project sequence counter |
| **D** | `DATE` | Recording date |
| **E** | `TITRE D'APPEL D'OFFRE` | Classified action (e.g., RÉALISATION, AMÉNAGEMENT) |
| **F** | `Nombre de projet` | Default `1` |
| **G** | `TYPE DE PROJET` | Facility type (e.g., STADE DE FOOTBALL, AIRE DE JEUX) |
| **H** | `DATE DE PARUTION` | Publication date |
| **I** | `DATE D'ECHEANCE` | Deadline date |
| **J** | `ANNONCEUR` | Contracting public entity |
| **K** | `WILAYA` | Algerian Wilaya |
| **L** | `COMMUNE` | Algerian Commune |
| **M** | `CATEGORIE` | Default `TRAVAUX PUBLICS` |
| **N** | `SCANS` | Direct clickable hyperlink to scan gallery on GitHub |

### 2. `AVIS D'ATTRIBUTIONS 2026`
| Col | Header | Description |
| :---: | :--- | :--- |
| **A** | `N°` | Incrementing project sequence counter |
| **B** | `DATE` | Recording date |
| **C** | `TITRE D'AVIS D'ATTRIBUTION` | Classified action (e.g., RÉALISATION, REVÊTEMENT) |
| **D** | `Nombre de projet` | Default `1` |
| **E** | `TYPE DE PROJET` | Facility type |
| **F** | `DATE DE PARUTION` | Original tender date published in newspapers (or `/` if consultation) |
| **G** | `DATE D'ECHEANCE` | Legal contestation deadline date |
| **H** | `ANNONCEUR` | Contracting public entity |
| **I** | `WILAYA` | Algerian Wilaya |
| **J** | `COMMUNE` | Algerian Commune |
| **K** | `DATE D'ATTRIBUTION` | Attribution decision date |
| **L** | `BUDGET` | Corrected winning amount TTC in Dinar (DA) |
| **M** | `ATTRIBUTION PAR` | Winning contractor / enterprise |
| **N** | `DELAI` | Execution delay (e.g., `03 MOIS`, `60 JOURS`) |
| **O** | `WILAYA 2` | Complementary territory info (`/`) |
| **P** | `SCANS` | Direct clickable hyperlink to scan gallery on GitHub |

---

## 🔒 Security & Best Practices

* Never commit `.env` or raw Google Service Account private keys (`state/service_account.json`).
* Authentication cookies in `state/cookies.json` are maintained automatically by GitHub Actions to keep web sessions warm.
