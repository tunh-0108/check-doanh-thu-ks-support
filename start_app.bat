@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Dang khoi dong app...
"venv\Scripts\python.exe" -m streamlit run app.py
pause
