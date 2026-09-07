@echo off
title Excel ISIN Splitter & Filter
cd /d "%~dp0"
echo ===================================================================
echo   Starting Excel ISIN Filter & Splitter Application...
echo ===================================================================
echo.
python -m streamlit run app.py
pause
