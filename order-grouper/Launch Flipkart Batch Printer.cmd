@echo off
cd /d "%~dp0"
if exist "%~dp0dist\FlipkartBatchPrinter.exe" (
    start "" "%~dp0dist\FlipkartBatchPrinter.exe"
) else (
    echo FlipkartBatchPrinter.exe was not found.
    echo Run build_exe.ps1 first.
    pause
)
