@echo off
cd /d "%~dp0"
if not exist .venv\Scripts\python.exe (
    py -m venv .venv
    if errorlevel 1 goto failed
)
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto failed
.venv\Scripts\python.exe import_firefox_tabs.py %*
if errorlevel 1 goto failed
pause
exit /b 0
:failed
echo Firefox tab import failed. See the error above.
pause
exit /b 1
