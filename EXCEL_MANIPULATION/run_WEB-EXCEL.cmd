@echo off
cd /d "%~dp0WEB-EXCEL"
call npm run dev -- --host 127.0.0.1 --port 5174
