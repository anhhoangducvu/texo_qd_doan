@echo off
title TEXO - Lap Quyet dinh Doan Tu van
cd /d "%~dp0"
echo ======================================================
echo 🏗️ TEXO - KHOI CHAY UNG DUNG LAP QUYET DINH
echo ======================================================
echo.
echo [1/2] Dang kiem tra va cai dat thu vien (neu can)...
pip install -r requirements.txt
echo.
echo [2/2] Dang khoi chay Streamlit...
streamlit run app.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ❌ Co loi xay ra khi khoi chay ung dung.
    pause
)
