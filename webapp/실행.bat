@echo off
chcp 65001 >nul
cd /d "%~dp0"
python -c "import flask, openpyxl" 2>nul || pip install flask openpyxl
python app.py
pause
