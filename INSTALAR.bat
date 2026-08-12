@echo off
REM Prepara el entorno del Gestor de Cobranzas Wood Tools.
cd /d "%~dp0"
echo ============================================================
echo   Gestor de Cobranzas Wood Tools - instalacion
echo ============================================================
echo.

where python >nul 2>nul
if not %errorlevel%==0 (
    echo ERROR: no se encontro Python.
    echo Instalalo desde https://www.python.org/downloads/
    echo IMPORTANTE: tildar "Add Python to PATH" durante la instalacion.
    pause
    exit /b 1
)

echo [1/3] Creando el entorno virtual...
if not exist "venv" python -m venv venv

echo [2/3] Instalando las librerias...
call venv\Scripts\python.exe -m pip install --upgrade pip
call venv\Scripts\python.exe -m pip install -r requirements.txt
if not %errorlevel%==0 (
    echo.
    echo ERROR instalando las librerias. Revisa la conexion a internet.
    pause
    exit /b 1
)

echo [3/3] Preparando la configuracion...
if not exist "config.json" (
    copy "config.example.json" "config.json" >nul
    echo    Se creo config.json. Abri el programa y carga tus claves
    echo    en la pestana Configuracion.
)

echo.
echo ============================================================
echo   Listo. Abri el programa con ABRIR_PROGRAMA.bat
echo ============================================================
pause
