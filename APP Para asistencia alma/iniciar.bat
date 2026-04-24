@echo off
title Asistencia Alma - Cargando...
cd /d "%~dp0"

:: Verificar Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python no esta instalado o no esta en el PATH.
    echo Descargalo de https://www.python.org/downloads/
    pause
    exit /b
)

:: Instalar dependencias si no estan
pip show flask >nul 2>&1 || pip install flask
pip show openpyxl >nul 2>&1 || pip install openpyxl

echo.
echo ========================================
echo   Asistencia Alma - Iniciando...
echo   No cierres esta ventana.
echo ========================================
echo.

python app.py
pause
