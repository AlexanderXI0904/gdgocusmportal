@echo off
title GDGoC Email Sender

echo ==========================================
echo Installing Requirements...
echo ==========================================
echo.

:: Ensure we are running in the correct folder
cd /d "%~dp0"

:: Start the Python application
pip install -r requirements.txt