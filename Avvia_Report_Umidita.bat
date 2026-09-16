@echo off
title Generatore Report Umidita SCADA
cd /d "%~dp0"
where python >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    python Humidity_Report_Generator.py
) else (
    py Humidity_Report_Generator.py
)
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Si e' verificato un errore durante l'avvio del Generatore Report Umidita.
    pause
)
