@echo off
title Bot Agricola Santa Elisa
cd /d "%~dp0"

echo ============================================
echo   Bot Agricola Santa Elisa - Inicio
echo ============================================
echo.

REM Detectar Python 3.11 (intenta varios paths)
set PYTHON=
where py >nul 2>&1
if %errorlevel% equ 0 (
    py -3.11 --version >nul 2>&1
    if %errorlevel% equ 0 set PYTHON=py -3.11
)

if "%PYTHON%"=="" (
    if exist "%LOCALAPPDATA%\Python\bin\python3.11.exe" (
        set "PYTHON=%LOCALAPPDATA%\Python\bin\python3.11.exe"
    )
)

if "%PYTHON%"=="" (
    if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" (
        set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    )
)

if "%PYTHON%"=="" (
    where python >nul 2>&1
    if %errorlevel% equ 0 set PYTHON=python
)

if "%PYTHON%"=="" (
    echo [ERROR] Python 3.11 no encontrado.
    echo Instala Python 3.11 o ajusta la ruta en iniciar_bot.bat
    pause
    exit /b 1
)

echo Python detectado: %PYTHON%
%PYTHON% --version
echo.

REM Verificar dependencias
echo Verificando dependencias...
%PYTHON% -m pip install -q python-telegram-bot python-dotenv openpyxl anthropic python-dateutil >nul 2>&1

REM Verificar .env
if not exist ".env" (
    echo [ERROR] Archivo .env no encontrado.
    echo Crea un archivo .env con TELEGRAM_TOKEN y ANTHROPIC_API_KEY
    pause
    exit /b 1
)

echo.
echo Iniciando bot con WATCHDOG (reinicio automatico)... (Ctrl+C para detener)
echo Nota: si ya hay un watchdog corriendo (tarea AgricolaBotWatchdog), este saldra solo.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0watchdog_bot.ps1"
pause
