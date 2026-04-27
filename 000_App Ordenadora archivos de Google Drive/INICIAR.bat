@echo off
title Gestor de Sesiones — Hernán Gabriel Stagni
color 2F

echo.
echo  ============================================
echo   Gestor de Sesiones — Hernán Gabriel Stagni
echo  ============================================
echo.

cd /d "%~dp0"

echo  [1/5] Verificando Python...
python --version
if errorlevel 1 (
    echo.
    echo  ERROR: Python no encontrado.
    echo  Instala Python desde https://python.org
    echo  Asegurate de tildar "Add Python to PATH" durante la instalacion.
    echo.
    pause
    exit
)

echo  [2/5] Verificando Flask...
pip show flask >nul 2>&1
if errorlevel 1 (
    echo  Instalando Flask...
    pip install flask
    echo  Flask instalado.
) else (
    echo  Flask OK.
)

echo  [3/5] Verificando mutagen (duracion de MP3)...
pip show mutagen >nul 2>&1
if errorlevel 1 (
    echo  Instalando mutagen...
    pip install mutagen
) else (
    echo  Mutagen OK.
)

echo  [4/5] Verificando openpyxl (registro Excel)...
pip show openpyxl >nul 2>&1
if errorlevel 1 (
    echo  Instalando openpyxl...
    pip install openpyxl
) else (
    echo  Openpyxl OK.
)

echo  [5/5] Verificando win10toast (notificaciones)...
pip show win10toast >nul 2>&1
if errorlevel 1 (
    echo  Instalando win10toast...
    pip install win10toast
) else (
    echo  Win10toast OK.
)

echo.
echo  ============================================
echo   Todo listo. Iniciando servidor...
echo   El navegador se abrira automaticamente.
echo.
echo   IMPORTANTE: Deja esta ventana abierta
echo   mientras uses el Gestor de Sesiones.
echo   Para cerrar el sistema, cierra esta ventana.
echo  ============================================
echo.

python app.py

echo.
echo  El servidor se detuvo.
pause
