@echo off
echo ==========================================
echo EJECUTANDO SISTEMA
echo ==========================================
echo.

REM Verificar si existe el entorno virtual
if not exist ".venv" (
    echo [ERROR] Entorno virtual no encontrado
    echo Por favor ejecuta primero: setup.bat
    pause
    exit /b 1
)

REM Activar entorno virtual y ejecutar
call .venv\Scripts\activate.bat
python system_complete_fixed.py

echo.
echo ==========================================
echo Programa finalizado
echo ==========================================
pause
