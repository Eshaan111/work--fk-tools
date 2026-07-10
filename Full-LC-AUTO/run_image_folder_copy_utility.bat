@echo off
cd /d "%~dp0"
python image_folder_copy_utility.py
if errorlevel 1 pause
