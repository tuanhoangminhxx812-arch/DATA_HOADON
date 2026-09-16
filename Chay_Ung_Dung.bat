@echo off
title Ung Dung Xu Ly Hoa Don MTMN
chcp 65001 > nul
cd /d "%~dp0"

echo ========================================================
echo   KHOI DONG UNG DUNG XU LY HOA DON DIEN TU MTMN
echo ========================================================
echo.
echo Dang mo ung dung tren trinh duyet...
echo.

python -m streamlit run app.py

pause
