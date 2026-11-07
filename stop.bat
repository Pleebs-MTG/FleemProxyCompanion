@echo off
REM Stops Fleem Proxy Companion (you can also click its icon in the taskbar corner and choose Quit).
taskkill /im "Fleem Proxy Companion.exe" /f >nul 2>nul && echo Stopped Fleem Proxy Companion.
powershell -NoProfile -Command "Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like '*fleem_pdf_maker.py*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force; 'Stopped.' }"
pause
