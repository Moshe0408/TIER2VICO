@echo off
chcp 65001 > nul

REM ════════════════════════════════════════════
REM  הכנס כאן את מפתח ה-API שלך מ-console.anthropic.com
REM  לדוגמה: set ANTHROPIC_API_KEY=sk-ant-api03-xxxxx
REM ════════════════════════════════════════════
set ANTHROPIC_API_KEY=YOUR_API_KEY_HERE

echo.
echo  ================================================================
echo   MosheAI - AI Agent for Reports, Slides and Statistics
echo  ================================================================
echo   משתמש: Moshei1   סיסמה: Admin2026
echo   כתובת: http://localhost:5000
echo   לסגירה: Ctrl+C
echo  ================================================================
echo.

start "" http://localhost:5000
python app.py

pause
