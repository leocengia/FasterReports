@echo off
setlocal enabledelayedexpansion
title FasterReports - Agent Performance Score

echo ============================================================
echo   FasterReports ^| Agent Performance Score
echo ============================================================
echo.
echo   Input atteso: export AHT blended (CSV separato da ;)
echo   Colonne necessarie: Employee Name, Case Channel,
echo                       Case AHT (mins), Distinct Cases
echo.
echo ------------------------------------------------------------
echo.

set /p INPUT_FILE=  Percorso file CSV raw (es. raw_data\AHT blended WEEK 24.csv):

if not exist "%INPUT_FILE%" (
    echo.
    echo   ERRORE: file non trovato - "%INPUT_FILE%"
    echo.
    pause
    exit /b 1
)

set /p WEEK=  Numero settimana (es. 24):

set DAYS=
set /p DAYS=  Giorni lavorativi nel mese [Invio = 22]:
if "!DAYS!"=="" set DAYS=22

echo.
echo   Generazione report in corso...
echo.

python src\agent_performance_score.py "%INPUT_FILE%" --week %WEEK% --days %DAYS%

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo   ERRORE durante la generazione. Controlla i messaggi sopra.
) else (
    echo.
    echo   Report salvato in: output\agent_performance_score_W%WEEK%.xlsx
)

echo.
pause
