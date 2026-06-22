@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0.."
title FasterReports - Case KPI Report

echo ============================================================
echo   FasterReports ^| Case KPI Report
echo ============================================================
echo.
echo   Cartella raw data: %cd%\raw_data
echo.
echo   File disponibili:
for %%f in ("raw_data\*.csv") do echo     - %%~nxf
echo.
echo   Input atteso: export completo casi (CSV separato da ;)
echo   Colonne necessarie: Case Type, Case AHT (mins), Distinct Cases
echo   Colonne opzionali:  Primary Category, Case Origin (group)
echo.
echo ------------------------------------------------------------
echo.

set /p FILENAME=  Nome file CSV (es. cases_W24.csv):

set INPUT_FILE=raw_data\!FILENAME!

if not exist "!INPUT_FILE!" (
    echo.
    echo   ERRORE: file non trovato - "!INPUT_FILE!"
    echo.
    pause
    exit /b 1
)

set /p WEEK=  Numero settimana (es. 24):

set TARGET_PHONE=
set /p TARGET_PHONE=  Target AHT Phone    [Invio = 19.98]:
if "!TARGET_PHONE!"=="" set TARGET_PHONE=19.98

set TARGET_NL=
set /p TARGET_NL=  Target AHT Non-live [Invio = 18.96]:
if "!TARGET_NL!"=="" set TARGET_NL=18.96

set CHANNEL=
set /p CHANNEL=  Filtro canale (Phone / Non-live / Invio = tutti):

echo.
echo   Generazione report in corso...
echo.

if "!CHANNEL!"=="" (
    python src\case_KPI_report.py "!INPUT_FILE!" --week !WEEK! --target-phone !TARGET_PHONE! --target-nonlive !TARGET_NL!
) else (
    python src\case_KPI_report.py "!INPUT_FILE!" --week !WEEK! --target-phone !TARGET_PHONE! --target-nonlive !TARGET_NL! --channel "!CHANNEL!"
)

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo   ERRORE durante la generazione. Controlla i messaggi sopra.
) else (
    echo.
    echo   Report salvato in: output\case_KPI_report_W!WEEK!.xlsx
)

echo.
pause
