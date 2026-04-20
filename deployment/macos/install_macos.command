#!/bin/sh
set -eu

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

mkdir -p "$INSTALL_ROOT"
mkdir -p "$APPLICATIONS_DIR"

rm -rf "$INSTALL_DIR"
ditto "$PACKAGE_APP_DIR" "$INSTALL_DIR"
chmod +x "$LAUNCHER_PATH"

cat > "$SHORTCUT_PATH" <<EOF
#!/bin/sh
cd "$INSTALL_DIR"
exec "$LAUNCHER_PATH" >/dev/null 2>&1 &
EOF
chmod +x "$SHORTCUT_PATH"

open "$SHORTCUT_PATH"

