@echo off
setlocal
cd /d "%~dp0"
call "..\.venv\Scripts\activate.bat"
python -m uvicorn rate_insight_web.app:app --host 127.0.0.1 --port 8000 --reload
