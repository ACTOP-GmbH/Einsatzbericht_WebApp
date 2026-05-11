@echo off
setlocal
title Einsatzbericht Manager Installation
set "SCRIPT_DIR=%~dp0"
echo Einsatzbericht Manager wird installiert oder aktualisiert.
echo Bitte dieses Fenster nicht schliessen, bis die Abschlussmeldung erscheint.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%install_windows.ps1"
if errorlevel 1 (
  echo.
  echo Installation failed.
  pause
  exit /b 1
)
echo.
echo Installation abgeschlossen.
timeout /t 2 /nobreak >nul
exit /b 0
