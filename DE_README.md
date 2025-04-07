# CSV zu XLS Konverter

Dieses Skript konvertiert CSV-Dateien (Comma Separated Values) in das XLS-Format (Excel 97-2003).

Es kann in zwei Modi betrieben werden:

1.  **Ordner-Modus:** Konvertiert alle `.csv`-Dateien, die in einem angegebenen Eingabeverzeichnis (Standard: `input/`) gefunden werden, und speichert die entsprechenden `.xls`-Dateien in einem Ausgabeverzeichnis (Standard: `output/`).
2.  **Einzeldatei-Modus:** Konvertiert eine bestimmte Eingabe-CSV-Datei in eine angegebene Ausgabe-XLS-Datei.

Das Skript versucht automatisch zu erkennen, ob das CSV-Trennzeichen ein Komma (`,`) oder ein Semikolon (`;`) ist.

## Projektstruktur

```
.
├── src/                   # Enthält die Hauptkonvertierungslogik
│   └── converter.py
├── input/                 # Standardverzeichnis für Eingabe-CSV-Dateien
├── output/                # Standardverzeichnis für Ausgabe-XLS-Dateien (wird bei Bedarf erstellt)
├── venv/                  # Beispiel für eine virtuelle Python-Umgebung (von Git ignoriert)
├── .gitignore             # Gibt Dateien/Ordner an, die Git ignorieren soll
├── README.md              # Englische Readme-Datei
├── DE_README.md           # Diese Datei (Deutsche Readme)
├── requirements.txt       # Listet Abhängigkeiten auf (xlwt)
├── setup_and_run.bat      # Windows Batch-Skript für Setup und Ausführung
└── start_conversion.py    # Einfache Möglichkeit, den Konverter vom Stammverzeichnis auszuführen
```

## Einrichtung

1.  **Repository klonen (falls zutreffend):**
    ```bash
    git clone <repository-url>
    cd <repository-ordner>
    ```

2.  **Virtuelle Umgebung erstellen (Empfohlen):**
    ```bash
    python -m venv venv
    ```
    *   Unter Windows: `venv\Scripts\activate`
    *   Unter macOS/Linux: `source venv/bin/activate`

3.  **Abhängigkeiten installieren:**
    Das Skript benötigt die `xlwt`-Bibliothek. Sie sollten diese über die `requirements.txt`-Datei installieren.
    ```bash
    pip install -r requirements.txt
    ```
    *(Alternativ manuelle Installation: `pip install xlwt`)*

4.  **Eingabedateien vorbereiten:**
    *   Legen Sie Ihre `.csv`-Dateien in das `input/`-Verzeichnis (oder erstellen Sie es, falls es nicht existiert).

## Verwendung

Sie können die Konvertierung mit dem Skript `start_conversion.py` vom Stammverzeichnis des Projekts aus starten oder unter Windows das Skript `setup_and_run.bat` verwenden, das die Einrichtung und Ausführung übernimmt.

**Verwendung von `setup_and_run.bat` (Windows):**

*   Doppelklicken Sie auf die Datei `setup_and_run.bat`.
*   Das Skript versucht, eine virtuelle Umgebung zu erstellen (falls nötig), installiert die Abhängigkeiten aus `requirements.txt` und führt dann den Konverter im Ordner-Modus aus.
*   Folgen Sie den Anweisungen im Konsolenfenster. Administratorrechte könnten erforderlich sein, falls eine Python-Installation über winget versucht wird.

**Verwendung von `start_conversion.py` (Manuell):**

Stellen Sie sicher, dass Sie Ihre virtuelle Umgebung aktiviert und die Abhängigkeiten installiert haben.

**1. Ordner-Modus (Standard):**

*   Stellen Sie sicher, dass sich Ihre CSV-Dateien im `input/`-Verzeichnis befinden.
*   Führen Sie das Start-Skript aus:
    ```bash
    python start_conversion.py
    ```
*   Das Skript verarbeitet alle `.csv`-Dateien in `input/` und speichert die resultierenden `.xls`-Dateien im `output/`-Verzeichnis.

**2. Einzeldatei-Modus:**

*   Verwenden Sie die Option `-i` (oder `--input`) für den Pfad der Eingabe-CSV-Datei und die Option `-o` (oder `--output`) für den gewünschten Pfad der Ausgabe-XLS-Datei.
    ```bash
    python start_conversion.py -i pfad/zu/ihrer/datei.csv -o pfad/zu/ihrer/ausgabe.xls
    ```
    *Beispiel mit Standardordnern:*
    ```bash
    python start_conversion.py -i input/spezifische_datei.csv -o output/konvertierte_spezifische_datei.xls
    ```

## Hinweise

*   Das Skript geht davon aus, dass die Eingabe-CSV-Dateien UTF-8-kodiert sind.
*   Das `output/`-Verzeichnis wird automatisch erstellt, falls es im Ordner-Modus nicht existiert.
*   Fehlermeldungen und Fortschrittsinformationen werden in der Konsole ausgegeben.
*   Das Ausgabeformat ist das ältere `.xls` (Excel 97-2003). Dieses Format hat Einschränkungen, z. B. maximal 65.536 Zeilen pro Blatt.
