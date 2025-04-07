@echo off
chcp 1252 > nul
setlocal enabledelayedexpansion

REM ============================================================================
REM Script: setup_and_run.bat
REM Purpose: Attempts to install Python (via winget), create venv, install
REM          requirements, and run the CSV to XLS converter.
REM REQUIRES: Administrator privileges, winget, Internet connection.
REM ============================================================================

title CSV zu XLS Konverter Setup und Ausführung



REM --- Check for winget ---
echo Überprüfe auf winget Paketmanager...
winget --version > nul 2>&1
if %errorLevel% == 0 (
    echo Winget gefunden.
) else (
    echo ============================================================
    echo "FEHLER: Windows Paketmanager (winget) nicht gefunden oder funktioniert nicht."
    echo Winget wird benötigt, um Python automatisch zu installieren.
    echo Bitte stellen Sie sicher, dass Sie eine moderne Windows-Version verwenden oder installieren Sie winget manuell.
    echo ============================================================
    pause
    exit /b 1
)



REM --- Check for requirements.txt ---
echo Überprüfe auf requirements.txt...
dir /b "requirements.txt" > nul 2>&1
if %errorLevel% == 0 (
    echo requirements.txt gefunden.
) else (
    echo ============================================================
    echo FEHLER: requirements.txt wurde im aktuellen Verzeichnis nicht gefunden.
    echo Diese Datei wird zur Installation der Abhängigkeiten benötigt.
    echo ============================================================
    pause
    exit /b 1
)

REM --- Check for Virtual Environment ---
echo Überprüfe auf virtuelle Umgebung (venv)...
dir /b ".\venv\Scripts\activate.bat" > nul 2>&1
if %errorLevel% == 0 (
    echo Virtuelle Umgebung gefunden.
) else (
    echo Virtuelle Umgebung nicht gefunden. Erstelle eine...
    python -m venv venv
    if %errorLevel% == 0 (
        echo Virtuelle Umgebung erfolgreich erstellt.
    ) else (
        echo ============================================================
        echo FEHLER: Konnte die virtuelle Umgebung nicht erstellen. Exit code: %errorLevel%
        echo Überprüfe die Python-Installation und Berechtigungen.
        echo ============================================================
        pause
        exit /b 1
    )
)

REM --- Activate Virtual Environment and Install Requirements ---
echo Aktiviere virtuelle Umgebung...
call .\venv\Scripts\activate.bat

echo Installiere/Aktualisiere Pakete aus requirements.txt...
pip install -r requirements.txt
if %errorLevel% == 0 (
    echo Pakete erfolgreich installiert/aktualisiert.
) else (
    echo ============================================================
    echo FEHLER: Konnte Pakete mit pip nicht installieren. Exit code: %errorLevel%
    echo Überprüfe die Internetverbindung und mögliche Fehlermeldungen oben.
    echo ============================================================
    pause
    exit /b 1
)

REM --- Check for start_conversion.py ---
echo Überprüfe auf start_conversion.py...
dir /b "start_conversion.py" > nul 2>&1
if %errorLevel% == 0 (
    echo start_conversion.py gefunden.
) else (
    echo ============================================================
    echo FEHLER: start_conversion.py wurde im aktuellen Verzeichnis nicht gefunden.
    echo ============================================================
    pause
    exit /b 1
)

REM --- Run the Conversion Script ---
echo Starte den CSV zu XLS Konvertierungsprozess...
echo Die Ausgabe des Konverters wird unten angezeigt:
echo --------------------------------------------
python start_conversion.py %*
REM (%* übergibt alle Argumente von setup_and_run.bat an start_conversion.py)

echo --------------------------------------------
echo Konvertierungsskript beendet.
pause
exit /b 0
