@echo off
title MarketRytrAI Automation Dashboard
echo ============================================
echo  MarketRytrAI Automation Dashboard Setup
echo ============================================
echo.

:: Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH.
    echo Please install Python 3.10+ from https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/4] Creating virtual environment...
if not exist "venv" (
    python -m venv venv
)

echo [2/4] Activating virtual environment...
call venv\Scripts\activate.bat

echo [3/4] Installing dependencies...
pip install -r requirements.txt --quiet
python -m playwright install chromium

echo [4/4] Starting application...
echo.
echo ============================================
echo  App running at: http://localhost:8501
echo  Press Ctrl+C to stop
echo ============================================
echo.
streamlit run app.py --server.port=8501

pause
