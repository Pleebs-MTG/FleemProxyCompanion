@echo off
REM Starts Fleem Proxy Companion (opens its window, or its icon by the clock if it runs in the background).
cd /d "%~dp0"
if exist "%~dp0Fleem Proxy Companion.exe" (start "" "%~dp0Fleem Proxy Companion.exe" & goto :eof)
where pyw >nul 2>nul && (start "" pyw "%~dp0fleem_pdf_maker.py" & goto :eof)
where pythonw >nul 2>nul && (start "" pythonw "%~dp0fleem_pdf_maker.py" & goto :eof)
echo Couldn't find "Fleem Proxy Companion.exe" or Python. Run build.bat first (or download the .exe from Releases).
pause
