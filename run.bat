@echo off
REM ── AlgerieMarches Scraper launcher ──
REM This file is called by Windows Task Scheduler every 6 hours.
cd /d "%~dp0"
call .venv\Scripts\activate.bat
python scraper.py
