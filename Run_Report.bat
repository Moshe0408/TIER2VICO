@echo off
chcp 65001 > nul
echo ===========================================
echo      VERINT REPORTER - OUTLOOK EDITION
echo ===========================================
echo.
echo [*] מפעיל בדיקת מיילים ועיבוד דוחות...
py Combined_Reporter.py

echo.
echo ===========================================
echo      התהליך הסתיים
echo ===========================================
pause
