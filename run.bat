@echo off
chcp 65001
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
cd /d "C:\Users\atsuk\OneDrive\デスクトップ\tech.app\financial-dashboard"
streamlit run app.py
pause
