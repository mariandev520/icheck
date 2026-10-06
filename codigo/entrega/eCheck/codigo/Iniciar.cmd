@echo off
cd /d "%~dp0"
where py >nul 2>nul
if not errorlevel 1 (
    py -3 app.py
    if errorlevel 1 pause
    exit /b
)
where python >nul 2>nul
if not errorlevel 1 (
    python app.py
    if errorlevel 1 pause
    exit /b
)
echo No se encontro Python. Usa eCheck.exe o instala Python con Tcl/Tk.
pause
