@echo off
setlocal
set "SETUP_SCRIPT=%~dp0Setup-FlipkartPrinter.ps1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath powershell.exe -Verb RunAs -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-NoExit','-File','%SETUP_SCRIPT%')"
endlocal
