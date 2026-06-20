@echo off
setlocal enabledelayedexpansion
title FasterReports - WoW Case Type Deep Dive

echo ============================================================
echo   FasterReports ^| WoW Case Type Deep Dive
echo ============================================================
echo.
echo   Input atteso: export completo casi (CSV separato da ;)
echo   Colonne necessarie: Case Type, Case AHT (mins),
echo                       Distinct Cases, Case Origin (group)
echo   Colonne opzionali:  MonthKey (es. 2026-06-01),
echo                       OOT_flag (0/1, derivato automaticamente
echo                       se assente)
echo.
echo ------------------------------------------------------------
echo.

set /p INPUT_FILE=  Percorso file CSV raw:

if not exist "%INPUT_FILE%" (
    echo.
    echo   ERRORE: file non trovato - "%INPUT_FILE%"
    echo.
    pause
    exit /b 1
)

set MONTH=
set /p MONTH=  Mese da analizzare (es. 2026-06) [Invio = tutti i dati]:

set TARGET_PHONE=
set /p TARGET_PHONE=  Target AHT Phone    [Invio = 19.98]:
if "!TARGET_PHONE!"=="" set TARGET_PHONE=19.98

set TARGET_NL=
set /p TARGET_NL=  Target AHT Non-live [Invio = 18.96]:
if "!TARGET_NL!"=="" set TARGET_NL=18.96

set MIN_VOL=
set /p MIN_VOL=  Volume minimo classificazione [Invio = 20]:
if "!MIN_VOL!"=="" set MIN_VOL=20

echo.
echo   Generazione report in corso...
echo.

if "!MONTH!"=="" (
    python src\wow_casetype_deepdive.py "%INPUT_FILE%" --target-phone %TARGET_PHONE% --target-nonlive %TARGET_NL% --min-volume %MIN_VOL%
) else (
    python src\wow_casetype_deepdive.py "%INPUT_FILE%" --month "!MONTH!" --target-phone %TARGET_PHONE% --target-nonlive %TARGET_NL% --min-volume %MIN_VOL%
)

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo   ERRORE durante la generazione. Controlla i messaggi sopra.
) else (
    echo.
    if "!MONTH!"=="" (
        echo   Report salvato in: output\wow_CaseType_Deepdive_All.xlsx
    ) else (
        echo   Report salvato in: output\wow_CaseType_Deepdive_!MONTH!.xlsx
    )
)

echo.
pause
