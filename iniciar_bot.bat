@echo off
title Bot Agricola Santa Elisa
cd /d "%~dp0"

echo ============================================
echo   Bot Agricola Santa Elisa - Inicio
echo ============================================
echo.

REM Verificar Python
py -3.11 --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python 3.11 no encontrado. Ejecuta: py install 3.11
    pause
    exit /b 1
)

REM Verificar dependencias
echo Verificando dependencias...
py -3.11 -m pip install -q python-telegram-bot python-dotenv openpyxl anthropic python-dateutil >nul 2>&1

REM Verificar .env
if not exist ".env" (
    echo [ERROR] Archivo .env no encontrado.
    echo Crea un archivo .env con TELEGRAM_TOKEN y ANTHROPIC_API_KEY
    pause
    exit /b 1
)

echo.
echo Iniciando bot... (Ctrl+C para detener)
echo.
py -3.11 main.py
pause
