@echo off
setlocal

set "APP_DIR=C:\work-mom\Code-Tools\EXCEL_MANIPULATION"
set "VENV_ACTIVATE=C:\work-mom\Code-Tools\.venv\Scripts\activate.bat"

if not exist "%VENV_ACTIVATE%" (
    echo Virtual environment activate script not found:
    echo %VENV_ACTIVATE%
    pause
    exit /b 1
)

call "%VENV_ACTIVATE%"
cd /d "%APP_DIR%"
python "Rate_insight.py"

set "EXIT_CODE=%ERRORLEVEL%"
if not "%EXIT_CODE%"=="0" (
    echo.
    echo Script exited with code %EXIT_CODE%.
    pause
)

endlocal & exit /b %EXIT_CODE%
