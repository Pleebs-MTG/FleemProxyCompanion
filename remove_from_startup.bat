@echo off
REM Stops Fleem Proxy Companion from starting when you log in.
cd /d "%~dp0"
if exist "%~dp0Fleem Proxy Companion.exe" (
    "%~dp0Fleem Proxy Companion.exe" --remove-startup
) else (
    set PY=py
    where py >nul 2>nul || set PY=python
    call %%PY%% fleem_pdf_maker.py --remove-startup
)
pause
