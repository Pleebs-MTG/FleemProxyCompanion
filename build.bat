@echo off
REM Builds "Fleem Proxy Companion.exe" (with Fleem's icon) in this folder. Run it once, and again after updates.
cd /d "%~dp0"
title Building Fleem Proxy Companion
set PY=py
where py >nul 2>nul || set PY=python
%PY% --version >nul 2>nul || (
    echo Python isn't installed. Get it from https://www.python.org/downloads/ and tick "Add Python to PATH".
    pause
    exit /b 1
)
echo Installing what Fleem Proxy Companion needs (this can take a minute)...
%PY% -m pip install --upgrade -r requirements.txt pyinstaller || goto :failed
echo Installing Windows text recognition support (lets Fleem read card names off cards)...
%PY% -m pip install --upgrade winrt-runtime winrt-Windows.Foundation winrt-Windows.Foundation.Collections ^
    winrt-Windows.Media.Ocr winrt-Windows.Graphics.Imaging winrt-Windows.Storage.Streams winrt-Windows.Globalization ^
    || echo (Skipped - Fleem works without it, it just won't read card names.)

echo.
%PY% fleem_pdf_maker.py --write-version-info || goto :failed
for /f %%v in ('%PY% fleem_pdf_maker.py --version') do set VER=%%v
echo Building Fleem Proxy Companion %VER%...
REM A folder build (--onedir): "Fleem Proxy Companion.exe" plus an _internal folder. It starts much
REM faster than a one-file .exe, which has to unpack itself into a temporary folder every time it opens.
%PY% -m PyInstaller --noconfirm --clean --onedir --windowed --name "Fleem Proxy Companion" ^
    --icon fleem.ico --add-data "fleem_icon.png;." --add-data "fonts;fonts" --version-file version_info.txt ^
    --hidden-import pystray._win32 --hidden-import pypdf --hidden-import fleem_mpcfill ^
    --collect-all winrt ^
    --collect-all pypdfium2 --collect-all pypdfium2_raw --collect-all pypdfium2_cfg fleem_pdf_maker.py || goto :failed

echo Closing Fleem Proxy Companion / the old helper if they're running...
taskkill /im "Fleem Proxy Companion.exe" /f >nul 2>nul
powershell -NoProfile -Command "Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like '*fleem_pdf_maker.py*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }" >nul 2>nul
timeout /t 2 /nobreak >nul
REM the .exe and its _internal folder go next to the other files here (start.bat etc. use them)
copy /y "dist\Fleem Proxy Companion\Fleem Proxy Companion.exe" "Fleem Proxy Companion.exe" >nul || goto :failed
if exist _internal rmdir /s /q _internal
robocopy "dist\Fleem Proxy Companion\_internal" "_internal" /E /NFL /NDL /NJH /NJS /NP >nul
if errorlevel 8 goto :failed
REM the release .zip for GitHub (what the auto-updater downloads)
if not exist release mkdir release
powershell -NoProfile -Command "Compress-Archive -Force -Path 'dist\Fleem Proxy Companion\*' -DestinationPath 'release\Fleem Proxy Companion %VER%.zip'" || goto :failed
rmdir /s /q build >nul 2>nul
rmdir /s /q dist >nul 2>nul
del /q "Fleem Proxy Companion.spec" >nul 2>nul

echo.
echo Done! "Fleem Proxy Companion.exe" (with its _internal folder) is in this folder.
echo For a GitHub release, upload "release\Fleem Proxy Companion %VER%.zip".
echo Next: double-click add_to_startup.bat (it starts it now and at every login).
pause
exit /b 0

:failed
echo.
echo The build failed - scroll up to see the error.
pause
exit /b 1
