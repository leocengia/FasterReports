@echo off
chcp 65001 > nul
echo ============================================================
echo   KPI Report Generator
echo ============================================================
echo.

:: ── Input settimanale ────────────────────────────────────────────────────────
set /p WEEK=Numero settimana (es. 24):
set /p CSV=Path CSV (es. raw_data/AHT blended WEEK %WEEK%.csv):

if not exist "%CSV%" (
    echo ERRORE: file non trovato: %CSV%
    pause
    exit /b 1
)

:: ── Generazione report ───────────────────────────────────────────────────────
echo.
echo Generazione report in corso...
python src\case_kpi_report.py generate "%CSV%" --week %WEEK%

if errorlevel 1 (
    echo.
    echo ERRORE durante la generazione del report.
    pause
    exit /b 1
)

:: ── Drill-down opzionale ─────────────────────────────────────────────────────
echo.
set /p ADD_DRILL=Aggiungere drill-down xLori ^& Costa? (s/n):
if /i "%ADD_DRILL%"=="s" (
    echo.
    echo Case type disponibili:
    echo   - Rates ^& Inventory Changes
    echo   - Booking Information
    echo   - Room Type/Rate Plan
    echo   - Partner Central Access
    echo   - Content Update
    echo   - EVC
    echo   - Expedia Collect Invoice
    echo   - Hotel Collect Issue
    echo   - Refund Request
    echo   - Supplier Initiated Traveler Contact
    echo   (altri: digitare il nome esatto)
    echo.
    set /p DRILL_TYPE=Case type per drill-down:
    echo.
    python src\case_kpi_report.py drill ^
        --output "output\case_KPI_report_W%WEEK%.xlsx" ^
        --input "%CSV%" ^
        --type "%DRILL_TYPE%"
    if errorlevel 1 (
        echo ERRORE durante l'aggiunta del drill-down.
        pause
        exit /b 1
    )
)

:: ── Fine ─────────────────────────────────────────────────────────────────────
echo.
echo ============================================================
echo   Report salvato: output\case_KPI_report_W%WEEK%.xlsx
echo ============================================================
echo.
echo Aprire il file adesso? (s/n)
set /p OPEN=
if /i "%OPEN%"=="s" (
    start "" "output\case_KPI_report_W%WEEK%.xlsx"
)
echo.
pause
