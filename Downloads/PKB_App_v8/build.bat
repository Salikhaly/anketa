@echo off
cd /d "%~dp0"
python -m pip install -r requirements.txt
python -m PyInstaller --onefile --windowed --name Anketa_PKB app.py
pause
