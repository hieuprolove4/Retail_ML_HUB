@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ====================================================================
echo    RETAIL ML HUB - KHOI DONG WEB APP VA MO TRINH DUYET TU DONG
echo ====================================================================
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" run_app.py
) else (
    python run_app.py
)
pause
