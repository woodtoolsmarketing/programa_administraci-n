@echo off
REM Corre el banco de pruebas del Gestor de Cobranzas Wood Tools.
REM No usa internet, no gasta credito de la API y no manda ningun WhatsApp:
REM todo lo que el programa escribe va a una carpeta temporal del sistema.
cd /d "%~dp0"

set PY=python
if exist "venv\Scripts\python.exe" set PY=venv\Scripts\python.exe

echo ============================================================
echo   Gestor de Cobranzas Wood Tools - banco de pruebas
echo ============================================================
echo.

%PY% tests\test_cobranzas.py
set RESULTADO=%errorlevel%

echo.
if %RESULTADO%==0 (
    echo ============================================================
    echo   TODO BIEN: no hay fallas.
    echo ============================================================
) else (
    echo ============================================================
    echo   HAY FALLAS. Mira el detalle de arriba.
    echo ============================================================
)
pause
exit /b %RESULTADO%
