@echo off
title Proteger Credenciales - Agricola Santa Elisa
cd /d "%~dp0"

echo ============================================
echo   Proteccion de Credenciales
echo   Agricola Santa Elisa
echo ============================================
echo.
echo Este script migra tus claves del .env al
echo Windows Credential Manager (almacen seguro).
echo.
echo Despues puedes limpiar el .env para que
echo las claves no queden en texto plano.
echo.

pip install keyring -q 2>nul
python credential_manager.py
pause
