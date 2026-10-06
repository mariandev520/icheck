@echo off
setlocal
cd /d "%~dp0"
if exist ".build-env\Scripts\python.exe" (
    ".build-env\Scripts\python.exe" recompilar.py
    goto terminado
)
where py >nul 2>nul
if not errorlevel 1 (
    py -3 recompilar.py
    goto terminado
)
where python >nul 2>nul
if not errorlevel 1 (
    python recompilar.py
    goto terminado
)
echo No se encontro Python. Instala Python de 64 bits con Tcl/Tk para compilar.
pause
exit /b 1
:terminado
set "resultado=%errorlevel%"
echo.
if not "%resultado%"=="0" echo No se pudo completar la compilacion. Revisa el error mostrado arriba.
pause
exit /b %resultado%
