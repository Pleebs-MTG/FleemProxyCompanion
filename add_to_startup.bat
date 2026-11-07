@echo off
REM Makes Fleem Proxy Companion start every time you log in to Windows (running in the background), and
REM starts it now. (The same as "Start with Windows" in Fleem's window.)
cd /d "%~dp0"
if exist "%~dp0Fleem Proxy Companion.exe" (
    "%~dp0Fleem Proxy Companion.exe" --install-startup
) else (
    set PY=py
    where py >nul 2>nul || set PY=python
    call %%PY%% fleem_pdf_maker.py --install-startup || goto :failed
)
call "%~dp0start.bat"
echo.
echo Done! Fleem Proxy Companion is running (look for its icon next to the clock) and will start whenever you log in.
pause
exit /b 0
:failed
echo.
echo Something went wrong - is Python installed? (https://www.python.org/downloads/, tick "Add Python to PATH")
pause
exit /b 1
