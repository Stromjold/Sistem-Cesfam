@echo off
echo ==========================================
echo INSTALADOR AUTOMATICO DE DEPENDENCIAS
echo ==========================================
echo.

REM Verificar si Python está instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python no está instalado o no está en PATH
    echo Por favor instala Python desde https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/3] Python detectado correctamente
echo.

REM Crear entorno virtual si no existe
if not exist ".venv" (
    echo [2/3] Creando entorno virtual...
    python -m venv .venv
    echo    Entorno virtual creado
) else (
    echo [2/3] Entorno virtual ya existe
)
echo.

REM Activar entorno virtual e instalar dependencias
echo [3/3] Instalando dependencias...
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt

echo.
echo ==========================================
echo INSTALACION COMPLETADA CON EXITO
echo ==========================================
echo.
echo Para ejecutar el programa:
echo   1. Activa el entorno: .venv\Scripts\activate.bat
echo   2. Ejecuta: python system_complete_fixed.py
echo.
echo O simplemente ejecuta: ejecutar.bat
echo.
pause
