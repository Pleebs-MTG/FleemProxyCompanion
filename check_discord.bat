@echo off
REM Checks the Discord order channels once and shows what happens (needs discord_token in settings.ini).
cd /d "%~dp0"
set PY=py
where py >nul 2>nul || set PY=python
%PY% fleem_pdf_maker.py --once
pause
