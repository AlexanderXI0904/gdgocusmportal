@echo off
title GDGoC USM Email Sender

echo ==========================================
echo Starting GDGoC USM Email Sender...
echo ==========================================
echo.

:: Ensure we are running in the correct folder
cd /d "%~dp0"

:: Open the default web browser to the local Flask server
start http://localhost:5000

:: Start the Python application
python app.py

:: Keep the window open just in case there is a crash or error
pause