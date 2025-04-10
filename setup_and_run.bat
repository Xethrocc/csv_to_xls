@echo off
chcp 1252 > nul
setlocal enabledelayedexpansion

REM ============================================================================
REM Script: setup_and_run.bat
REM Purpose: Attempts to install Python (via winget), create venv, install
REM          requirements, and run the CSV to XLS converter.
REM REQUIRES: winget, Internet connection. Administrator privileges likely needed for Python installation.
REM ============================================================================

title CSV zu XLS Konverter Setup und Ausführung

REM --- Check for winget ---
echo "Überprüfe auf winget Paketmanager..."
winget --version > nul 2>&1
if %errorLevel% == 0 (
    echo "Winget gefunden."
) else (
    echo "============================================================"
    echo "FEHLER: Windows Paketmanager (winget) nicht gefunden oder funktioniert nicht."
    echo "Winget wird benötigt, um Python automatisch zu installieren."
    echo "Bitte stellen Sie sicher, dass Sie eine moderne Windows-Version verwenden oder installieren Sie winget manuell."
    echo "============================================================"
    pause
    exit /b 1
)

REM --- Check for requirements.txt ---
echo "Überprüfe auf requirements.txt..."
dir /b "requirements.txt" > nul 2>&1
if %errorLevel% == 0 (
    echo "requirements.txt gefunden."
) else (
    echo "============================================================"
    echo "FEHLER: requirements.txt wurde im aktuellen Verzeichnis nicht gefunden."
    echo "Diese Datei wird zur Installation der Abhängigkeiten benötigt."
    echo "============================================================"
    pause
    exit /b 1
)

REM --- Check for Python Installation ---
REM We need Python to create the virtual environment.
echo "Überprüfe auf vorhandene Python-Installation..."
python --version > nul 2>&1
if %errorLevel% == 0 (
    echo "Python gefunden."
) else (
    echo "Python nicht im PATH gefunden. Versuche Installation mit winget..."
    echo "HINWEIS: Die Installation von Python erfordert möglicherweise Administratorrechte."
    winget install --id Python.Python.3 --exact --accept-package-agreements --accept-source-agreements
    if %errorLevel% == 0 (
        echo "Python-Installation (möglicherweise) erfolgreich."
        REM Re-check if python is now available in this session's PATH
        python --version > nul 2>&1
        if not %errorLevel% == 0 (
             echo "============================================================"
             echo "WARNUNG: Python wurde möglicherweise installiert, aber der 'python'-Befehl ist immer noch nicht im PATH dieser Sitzung."
             echo "Das Skript wird versuchen fortzufahren, aber die venv-Erstellung könnte fehlschlagen."
             echo "Möglicherweise müssen Sie dieses Fenster schließen und das Skript erneut starten oder Python manuell zum PATH hinzufügen."
             echo "============================================================"
             pause
        )
    ) else (
        echo "============================================================"
        echo "FEHLER: winget konnte Python nicht installieren (Exit Code: %errorLevel%)."
        echo "Bitte installiere Python 3 manuell und stelle sicher, dass es zum PATH hinzugefügt wird."
        echo "============================================================"
        pause
        exit /b 1
    )
)

REM --- Ensure Python is usable before creating venv ---
python --version > nul 2>&1
if not %errorLevel% == 0 (
    echo "============================================================"
    echo "FEHLER: Python ist nicht verfügbar, obwohl es vorhanden sein sollte oder gerade installiert wurde. Kann venv nicht erstellen."
    echo "Überprüfe die Python-Installation und PATH-Variable. Eventuell ist ein Neustart des Terminals nötig."
    echo "============================================================"
    pause
    exit /b 1
)

REM --- Check for Virtual Environment ---
echo "Überprüfe auf virtuelle Umgebung (venv)..."
dir /b ".\venv\Scripts\activate.bat" > nul 2>&1
if %errorLevel% == 0 (
    echo "Virtuelle Umgebung gefunden."
) else (
    echo "Virtuelle Umgebung nicht gefunden. Erstelle eine..."
    python -m venv venv
    if %errorLevel% == 0 (
        echo "Virtuelle Umgebung erfolgreich erstellt."
    ) else (
        echo "============================================================"
        echo "FEHLER: Konnte die virtuelle Umgebung nicht erstellen. Exit code: %errorLevel%"
        echo "Überprüfe die Python-Installation und Berechtigungen."
        echo "============================================================"
        pause
        exit /b 1
    )
)

REM --- Activate Virtual Environment and Install Requirements ---
echo "Aktiviere virtuelle Umgebung..."
call .\venv\Scripts\activate.bat

echo "Installiere/Aktualisiere Pakete aus requirements.txt..."
pip install -r requirements.txt
if %errorLevel% == 0 (
    echo "Pakete erfolgreich installiert/aktualisiert."
) else (
    echo "============================================================"
    echo "FEHLER: Konnte Pakete mit pip nicht installieren. Exit code: %errorLevel%"
    echo "Überprüfe die Internetverbindung und mögliche Fehlermeldungen oben."
    echo "============================================================"
    pause
    exit /b 1
)

REM --- Check for start_conversion.py ---
echo "Überprüfe auf start_conversion.py..."
dir /b "start_conversion.py" > nul 2>&1
if %errorLevel% == 0 (
    echo "start_conversion.py gefunden."
) else (
    echo "============================================================"
    echo "FEHLER: start_conversion.py wurde im aktuellen Verzeichnis nicht gefunden."
    echo "============================================================"
    pause
    exit /b 1
)

REM --- Run the Conversion Script ---
echo "Starte den CSV zu XLS Konvertierungsprozess..."
echo "Die Ausgabe des Konverters wird unten angezeigt:"
echo "--------------------------------------------"
python start_conversion.py %*
REM (%* übergibt alle Argumente von setup_and_run.bat an start_conversion.py)

echo "--------------------------------------------"
echo "Konvertierungsskript beendet."
pause
exit /b 0
