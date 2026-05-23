# Release-Build Anleitung

Diese Anleitung beschreibt den Windows-Release-Build fuer den Einsatzbericht Manager.

## Ergebnis

Die Produktivdatei fuer Windows ist:

```bat
release\EinsatzberichtManager-windows.zip
```

Diese ZIP wird in GitHub Releases hochgeladen. Tester starten daraus `install.bat`.

Die ZIP muss beim Entpacken diese Struktur haben:

```text
app\
install.bat
install_windows.ps1
README_INSTALL.txt
```

Im Ordner `app\` liegt dann der eigentliche PyInstaller-Payload:

```text
app\_internal\
app\data\
app\release_manifest.json
app\run_app.exe
app\streamlit_einsatzbericht_app_v2_excel_masterdata.py
```

Nicht direkt `dist\run_app` zippen. Dieser Ordner ist nur das PyInstaller-Zwischenprodukt und enthaelt typischerweise nur:

```text
_internal\
run_app.exe
```

## Voraussetzungen

Alle Befehle aus dem Repo-Root ausfuehren.

In CMD:

```bat
cd /d D:\Programs\einsatzbericht_streamlit_mvp_v2_print_layout
```

In PowerShell:

```powershell
cd D:\Programs\einsatzbericht_streamlit_mvp_v2_print_layout
```

Python pruefen:

```bat
python --version
```

Die lokale venv ist nicht zwingend erforderlich. Entscheidend ist:

- Der Python, der PyInstaller startet, bestimmt, welche Pakete in die EXE gebuendelt werden.
- Wenn dein globales `python -m PyInstaller ...` bisher funktioniert hat, kannst du diesen Weg weiter nutzen.
- Die `.venv` ist nur der reproduzierbarere Weg, weil die Abhaengigkeiten dann projektlokal festliegen.

Optional: Python aus der lokalen venv pruefen:

```bat
.\.venv\Scripts\python.exe --version
```

Falls beim Build diese Meldung kommt:

```text
No module named PyInstaller
```

dann fehlt PyInstaller in der lokalen `.venv`. Einmalig installieren:

```bat
.\.venv\Scripts\python.exe -m pip install pyinstaller
```

## Release bauen

### Variante A: Bisheriger CMD-Ablauf mit globalem PyInstaller

Diese Variante passt zu dem alten Ablauf, bei dem PyInstaller direkt vom PC/globalen Python kommt.

```bat
python -m PyInstaller run_app.spec --noconfirm
```

Danach das Release-ZIP aus dem fertigen `dist\run_app` bauen:

```bat
.\.venv\Scripts\python.exe scripts\build_desktop_release.py --platform windows --version v20260523_1
```

Wenn der globale Python auch `openpyxl` installiert hat, geht auch:

```bat
python scripts\build_desktop_release.py --platform windows --version v20260523_1
```

Wichtig: Beim zweiten Befehl hier **kein** `--build-pyinstaller` verwenden, weil PyInstaller bereits im ersten Schritt gelaufen ist.

Das Ergebnis ist danach nicht `dist\run_app`, sondern:

```bat
release\EinsatzberichtManager-windows.zip
```

Nach Aenderungen an `desktop_launcher.py`, `run_app.spec`, `desktop_runtime.py` oder der Streamlit-App immer zuerst PyInstaller neu ausfuehren. Sonst landet die alte EXE im Release-ZIP.

### Variante B: Ein-Schritt-Build ueber die venv

Empfohlener Ein-Schritt-Build fuer Windows:

```bat
.\.venv\Scripts\python.exe scripts\build_desktop_release.py --platform windows --build-pyinstaller --version v20260523_1
```

Diese Variante funktioniert nur, wenn PyInstaller in der `.venv` installiert ist. Falls nicht:

```bat
.\.venv\Scripts\python.exe -m pip install pyinstaller
```

Die Version am Ende anpassen, z. B. `v20260523_2` fuer einen zweiten Build am gleichen Tag.

Was der Befehl macht:

- baut `dist\run_app\run_app.exe` mit `run_app.spec`
- kopiert den Desktop-Payload in ein Release-Staging-Verzeichnis
- legt die Installer-Struktur unter `release\release_windows` an
- schreibt die Release-Version in die mitgelieferte `release_manifest.json`
- bereinigt die mitgelieferte Excel-Startdatei von Nutzerdaten
- erstellt `release\EinsatzberichtManager-windows.zip`

## Schneller Rebuild, wenn PyInstaller schon aktuell ist

Nur verwenden, wenn sich der PyInstaller-Payload nicht neu bauen muss:

```bat
.\.venv\Scripts\python.exe scripts\build_desktop_release.py --platform windows --version v20260523_1
```

## Ergebnis pruefen

Nach dem Build:

```bat
dir release\EinsatzberichtManager-windows.zip
dir release\release_windows
dir release\release_windows\app
```

Alternativ in PowerShell:

```powershell
Test-Path release\EinsatzberichtManager-windows.zip
Get-Item release\EinsatzberichtManager-windows.zip
```

Optional entpacken und kurz testen:

In PowerShell:

```powershell
Expand-Archive -Force release\EinsatzberichtManager-windows.zip release\_test_windows
release\_test_windows\install.bat
```

## GitHub Release

1. GitHub Release fuer die neue Version anlegen.
2. `release\EinsatzberichtManager-windows.zip` als Asset hochladen.
3. Release Notes in den GitHub Release Body schreiben.

Wichtig: Der Release Body wird in der App als Changelog angezeigt.

## Update-Hinweise

- `release_manifest.json` im Repo darf fuer lokale Entwicklung `version: "dev"` behalten.
- Das Build-Skript schreibt die echte Version in die Manifest-Datei innerhalb der Release-ZIP.
- Updates werden beim App-Start ueber das GitHub Release aus `release_manifest.json` geprueft.
- Nutzerdaten liegen ausserhalb des App-Payloads und werden durch Updates nicht ersetzt.

## Wenn etwas schiefgeht

Alte Build-Ausgaben koennen geloescht werden:

In CMD:

```bat
rmdir /s /q build
rmdir /s /q dist
rmdir /s /q release\release_windows
```

In PowerShell:

```powershell
Remove-Item -Recurse -Force build, dist, release\release_windows -ErrorAction SilentlyContinue
```

Danach den empfohlenen Ein-Schritt-Build erneut ausfuehren.
