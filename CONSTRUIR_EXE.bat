@echo off
REM Genera el .exe con PyInstaller (mismo esquema que WhatsAppMessage).
cd /d "%~dp0"

set PY=python
if exist "venv\Scripts\python.exe" set PY=venv\Scripts\python.exe

echo Instalando PyInstaller si falta...
%PY% -m pip install --upgrade pyinstaller

echo.
echo Construyendo el ejecutable...
REM --noconsole  : sin ventana negra (por eso el codigo nunca usa print)
REM --onefile    : un solo .exe
REM Los .json y el log quedan AL LADO del .exe (ruta_persistente en config.py),
REM asi sobreviven a las actualizaciones.
%PY% -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --noconsole ^
  --onefile ^
  --name "Gestor de Cobranzas Wood Tools" ^
  --collect-all reportlab ^
  --collect-all pdfplumber ^
  --hidden-import xlrd ^
  --hidden-import openpyxl ^
  --hidden-import anthropic ^
  interfaz.py

if not %errorlevel%==0 (
    echo.
    echo ERROR construyendo el ejecutable.
    pause
    exit /b 1
)

echo.
echo Copiando la configuracion de ejemplo junto al .exe...
if exist "dist" copy "config.example.json" "dist\config.example.json" >nul

echo.
echo ============================================================
echo   Listo: dist\Gestor de Cobranzas Wood Tools.exe
echo ============================================================
pause
