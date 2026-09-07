@echo off
title Generatore Report Umidita SCADA
cd /d "%~dp0"
python Humidity_Report_Generator.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Si e verificato un errore durante l'avvio del Generatore Report Umidita.
    pause
)
