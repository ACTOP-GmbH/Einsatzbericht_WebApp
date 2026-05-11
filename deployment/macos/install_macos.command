#!/bin/sh
set -eu

notify() {
  osascript -e "display notification \"$1\" with title \"Einsatzbericht Manager\"" >/dev/null 2>&1 || true
}

echo "Einsatzbericht Manager wird installiert oder aktualisiert."
echo "Bitte dieses Fenster nicht schliessen, bis die Abschlussmeldung erscheint."
notify "Installation wird gestartet."

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PACKAGE_APP_DIR="$SCRIPT_DIR/app"

if [ ! -d "$PACKAGE_APP_DIR" ]; then
  echo "Package payload not found: $PACKAGE_APP_DIR"
  exit 1
fi

APP_SLUG="actop-einsatzbericht-manager"
INSTALL_ROOT="$HOME/Library/Application Support/$APP_SLUG"
INSTALL_DIR="$INSTALL_ROOT/app"
LAUNCHER_PATH="$INSTALL_DIR/run_app"
APPLICATIONS_DIR="$HOME/Applications"
SHORTCUT_PATH="$APPLICATIONS_DIR/Einsatzbericht Manager.command"

echo ""
echo "[1/5] Installationsordner werden vorbereitet"
mkdir -p "$INSTALL_ROOT"
mkdir -p "$APPLICATIONS_DIR"
mkdir -p "$INSTALL_DIR"

echo "[2/5] Alte App-Dateien werden ersetzt"
if [ -d "$INSTALL_DIR/data" ]; then
  find "$INSTALL_DIR" -mindepth 1 -maxdepth 1 ! -name data -exec rm -rf {} +
  PRESERVE_EXISTING_DATA=1
else
  rm -rf "$INSTALL_DIR"
  mkdir -p "$INSTALL_DIR"
  PRESERVE_EXISTING_DATA=0
fi
echo "[3/5] App-Dateien werden kopiert"
if [ "$PRESERVE_EXISTING_DATA" -eq 1 ]; then
  for item in "$PACKAGE_APP_DIR"/* "$PACKAGE_APP_DIR"/.[!.]* "$PACKAGE_APP_DIR"/..?*; do
    [ -e "$item" ] || continue
    name="$(basename "$item")"
    [ "$name" = "data" ] && continue
    ditto "$item" "$INSTALL_DIR/$name"
  done
else
  ditto "$PACKAGE_APP_DIR" "$INSTALL_DIR"
fi
echo "[4/5] Starter wird vorbereitet"
chmod +x "$LAUNCHER_PATH"

cat > "$SHORTCUT_PATH" <<EOF
#!/bin/sh
cd "$INSTALL_DIR"
exec "$LAUNCHER_PATH" >/dev/null 2>&1 &
EOF
chmod +x "$SHORTCUT_PATH"

echo "[5/5] App wird gestartet"
open "$SHORTCUT_PATH"
echo ""
echo "Installation abgeschlossen."
notify "Installation abgeschlossen. Die App wurde gestartet."
osascript -e 'display dialog "Die App wurde erfolgreich installiert und gestartet." buttons {"OK"} default button "OK" with title "Installation abgeschlossen"' >/dev/null 2>&1 || true
