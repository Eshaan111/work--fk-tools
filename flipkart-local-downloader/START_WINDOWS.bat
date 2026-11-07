@echo off
cd /d "%~dp0"
if not exist .venv (py -m venv .venv)
call .venv\Scripts\activate.bat
python -m pip install -r requirements.txt
python -m playwright install chromium
python downloader.py --csv urls.csv --format jpg %*
pause
