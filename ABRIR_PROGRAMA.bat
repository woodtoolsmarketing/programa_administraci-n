@echo off
REM Abre el Gestor de Cobranzas Wood Tools en modo desarrollo.
cd /d "%~dp0"

if exist "venv\Scripts\pythonw.exe" (
    start "" "venv\Scripts\pythonw.exe" interfaz.py
    goto :eof
)

where pythonw >nul 2>nul
if %errorlevel%==0 (
    start "" pythonw interfaz.py
    goto :eof
)

echo No se encontro Python. Instalalo desde python.org y despues corre INSTALAR.bat
pause
