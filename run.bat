@echo off
chcp 65001 >nul
cd /d "%~dp0"

python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo 正在安装依赖 openpyxl...
    python -m pip install -r requirements.txt
)

python main.py
if errorlevel 1 pause
