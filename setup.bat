@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo ============================================================
echo   KPI Report Generator - Setup dipendenze
echo ============================================================
echo.

:: ── Verifica Python nel PATH ──────────────────────────────────────────────────
python --version > nul 2>&1
if errorlevel 1 (
    echo ERRORE: Python non trovato nel PATH.
    echo Installa Python 3.13 da https://www.python.org/downloads/
    echo e assicurati di spuntare "Add Python to PATH" durante l'installazione.
    echo.
    pause
    exit /b 1
)

for /f "tokens=*" %%v in ('python --version 2^>^&1') do echo Python rilevato: %%v

:: ── Crea venv se non esiste ───────────────────────────────────────────────────
if exist ".venv\Scripts\python.exe" (
    echo Ambiente virtuale .venv gia' presente, salto la creazione.
) else (
    echo.
    echo Creazione ambiente virtuale .venv ...
    python -m venv .venv
    if errorlevel 1 (
        echo ERRORE: impossibile creare il venv.
        pause
        exit /b 1
    )
    echo Ambiente virtuale creato.
)

:: ── Aggiorna pip ──────────────────────────────────────────────────────────────
echo.
echo Aggiornamento pip ...
".venv\Scripts\python.exe" -m pip install --upgrade pip --quiet
if errorlevel 1 (
    echo ERRORE durante l'aggiornamento di pip.
    pause
    exit /b 1
)

:: ── Installa dipendenze ───────────────────────────────────────────────────────
echo.
echo Installazione dipendenze da requirements.txt ...
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
    echo ERRORE durante l'installazione delle dipendenze.
    pause
    exit /b 1
)

:: ── Fine ──────────────────────────────────────────────────────────────────────
echo.
echo ============================================================
echo   Setup completato. Ora puoi lanciare run_weekly_report.bat
echo ============================================================
echo.
pause
