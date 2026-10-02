@echo off
chcp 65001 >nul
rem Перезбирає "PDF в XLSX.exe" у теці dist\
rem Необов'язковий brand.txt (підпис у шапці вікна) вшивається, якщо він є.
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  py -3.13 -m venv .venv
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt pyinstaller
)
set BRAND=
if exist brand.txt set BRAND=--add-data "brand.txt;."
".venv\Scripts\python.exe" make_icon.py
".venv\Scripts\pyinstaller.exe" --noconfirm --clean --onefile --windowed ^
  --name "PDF в XLSX" --icon icon.ico --add-data "icon.ico;." %BRAND% ^
  --collect-all tkinterdnd2 --exclude-module numpy ^
  pdf_to_xlsx.py
