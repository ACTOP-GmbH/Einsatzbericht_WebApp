from __future__ import annotations

import datetime as dt
import json
import os
import shutil
import subprocess
import sys
import tempfile
import textwrap
import urllib.request
from pathlib import Path
from typing import Any, Dict, Optional
import traceback


DEFAULT_MANIFEST: Dict[str, Any] = {
    "app_name": "Einsatzbericht Manager",
    "app_slug": "actop-einsatzbericht-manager",
    "github_repo": "ACTOP-GmbH/Einsatzbericht_WebApp",
    "version": "dev",
    "main_script": "streamlit_einsatzbericht_app_v2_excel_masterdata.py",
    "seed_workbook_relative_path": "data/Taetigkeiten_Ueberblick.xlsx",
    "seed_workbook_name": "Taetigkeiten_Ueberblick.xlsx",
    "release_asset_windows": "EinsatzberichtManager-windows.zip",
    "release_asset_macos": "EinsatzberichtManager-macos.zip",
    "launcher_relative_path_windows": "run_app.exe",
    "launcher_relative_path_macos": "run_app",
    "check_updates": True,
    "update_interval_hours": 12,
}


def _install_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _resource_dir() -> Path:
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", "")
        if meipass:
            return Path(meipass).resolve()
    return Path(__file__).resolve().parent


def load_release_manifest(base_dir: Optional[Path] = None) -> Dict[str, Any]:
    base_dir = base_dir or _resource_dir()
    manifest_path = base_dir / "release_manifest.json"
    manifest = dict(DEFAULT_MANIFEST)
    if manifest_path.exists():
        try:
            manifest.update(json.loads(manifest_path.read_text(encoding="utf-8")))
        except Exception:
            pass
    return manifest


def _user_data_root(manifest: Dict[str, Any]) -> Path:
    app_slug = str(manifest.get("app_slug") or DEFAULT_MANIFEST["app_slug"]).strip()
    if sys.platform == "win32":
        root = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    elif sys.platform == "darwin":
        root = Path.home() / "Library" / "Application Support"
    else:
        root = Path.home() / ".local" / "share"
    return root / app_slug


def _release_notes_path(user_root: Path) -> Path:
    target = user_root / "logs" / "update_state.json"
    target.parent.mkdir(parents=True, exist_ok=True)
    return target


def _seed_workbook_target(user_root: Path, manifest: Dict[str, Any]) -> Path:
    workbook_name = str(manifest.get("seed_workbook_name") or DEFAULT_MANIFEST["seed_workbook_name"])
    target_dir = user_root / "data"
    target_dir.mkdir(parents=True, exist_ok=True)
    return target_dir / workbook_name


def _seed_workbook_source(base_dir: Path, manifest: Dict[str, Any]) -> Path:
    rel_path = str(manifest.get("seed_workbook_relative_path") or DEFAULT_MANIFEST["seed_workbook_relative_path"])
    source = base_dir / rel_path
    if source.exists():
        return source
    for fallback_name in ["Tätigkeiten_Überblick.xlsx", "Taetigkeiten_Ueberblick.xlsx"]:
        fallback = base_dir / "data" / fallback_name
        if fallback.exists():
            return fallback
    return base_dir / "data" / "Taetigkeiten_Ueberblick.xlsx"


def prepare_runtime_environment(base_dir: Optional[Path] = None) -> Dict[str, Path]:
    resource_dir = base_dir or _resource_dir()
    install_dir = _install_dir()
    manifest = load_release_manifest(resource_dir)
    user_root = _user_data_root(manifest)
    for rel in ["data", "imports", "imports_reports", "logs", "downloads"]:
        (user_root / rel).mkdir(parents=True, exist_ok=True)

    seed_source = _seed_workbook_source(resource_dir, manifest)
    seed_target = _seed_workbook_target(user_root, manifest)
    if seed_source.exists() and not seed_target.exists():
        shutil.copy2(seed_source, seed_target)

    os.environ["EINSATZBERICHT_USER_DATA_DIR"] = str(user_root)
    os.environ["EINSATZBERICHT_DEFAULT_EXCEL"] = str(seed_target)

    return {
        "install_dir": install_dir,
        "resource_dir": resource_dir,
        "user_root": user_root,
        "default_excel": seed_target,
    }


def _load_update_state(user_root: Path) -> Dict[str, Any]:
    state_path = _release_notes_path(user_root)
    if not state_path.exists():
        return {}
    try:
        return json.loads(state_path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _save_update_state(user_root: Path, state: Dict[str, Any]) -> bool:
    state_path = _release_notes_path(user_root)
    try:
        state_path.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        return False
    return True


def _should_check_for_updates(user_root: Path, manifest: Dict[str, Any]) -> bool:
    if not bool(manifest.get("check_updates", True)):
        return False
    state = _load_update_state(user_root)
    interval_hours = int(manifest.get("update_interval_hours", 12) or 12)
    last_check = str(state.get("last_check_utc") or "").strip()
    if not last_check:
        return True
    try:
        last_dt = dt.datetime.fromisoformat(last_check)
    except Exception:
        return True
    return (dt.datetime.utcnow() - last_dt) >= dt.timedelta(hours=interval_hours)


def _fetch_latest_release(manifest: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    repo = str(manifest.get("github_repo") or "").strip()
    if not repo:
        return None
    url = f"https://api.github.com/repos/{repo}/releases/latest"
    req = urllib.request.Request(url, headers={"User-Agent": "einsatzbericht-manager-updater"})
    with urllib.request.urlopen(req, timeout=10) as response:
        return json.loads(response.read().decode("utf-8"))


def _release_asset_name(manifest: Dict[str, Any]) -> str:
    if sys.platform == "win32":
        return str(manifest.get("release_asset_windows") or DEFAULT_MANIFEST["release_asset_windows"])
    return str(manifest.get("release_asset_macos") or DEFAULT_MANIFEST["release_asset_macos"])


def _release_asset_url(release: Dict[str, Any], manifest: Dict[str, Any]) -> Optional[str]:
    expected_name = _release_asset_name(manifest)
    for asset in release.get("assets", []) or []:
        if str(asset.get("name") or "").strip() == expected_name:
            return str(asset.get("browser_download_url") or "").strip() or None
    return None


def _ask_yes_no(title: str, message: str) -> bool:
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        try:
            return bool(messagebox.askyesno(title, message))
        finally:
            root.destroy()
    except Exception:
        return False


def _show_error(title: str, message: str) -> None:
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        try:
            messagebox.showerror(title, message)
        finally:
            root.destroy()
    except Exception:
        pass


def report_startup_failure(exc: BaseException) -> Optional[Path]:
    try:
        runtime = prepare_runtime_environment()
        log_dir = runtime["user_root"] / "logs"
    except Exception:
        log_dir = _install_dir()
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / "startup_error.log"
    payload = "\n".join(
        [
            dt.datetime.now().isoformat(),
            "".join(traceback.format_exception(type(exc), exc, exc.__traceback__)),
            "-" * 80,
        ]
    )
    try:
        with log_path.open("a", encoding="utf-8") as handle:
            handle.write(payload + "\n")
    except Exception:
        log_path = None

    message = "Die App konnte nicht gestartet werden."
    if log_path is not None:
        message += f"\n\nDetails stehen in:\n{log_path}"
    _show_error("Startfehler", message)
    return log_path


def _download_release_asset(url: str, user_root: Path) -> Path:
    downloads_dir = user_root / "downloads"
    downloads_dir.mkdir(parents=True, exist_ok=True)
    file_name = Path(url.split("?")[0]).name or "update.zip"
    target = downloads_dir / file_name
    req = urllib.request.Request(url, headers={"User-Agent": "einsatzbericht-manager-updater"})
    with urllib.request.urlopen(req, timeout=30) as response:
        target.write_bytes(response.read())
    return target


def _write_windows_updater_script(base_dir: Path, zip_path: Path, manifest: Dict[str, Any]) -> Path:
    launcher_name = str(manifest.get("launcher_relative_path_windows") or DEFAULT_MANIFEST["launcher_relative_path_windows"])
    relaunch_path = base_dir / launcher_name
    script_path = Path(tempfile.gettempdir()) / "einsatzbericht_apply_update.ps1"
    payload = f"""
Start-Sleep -Seconds 2
$ZipPath = {json.dumps(str(zip_path))}
$InstallDir = {json.dumps(str(base_dir))}
$Relaunch = {json.dumps(str(relaunch_path))}
$Stage = Join-Path $env:TEMP ("einsatzbericht_update_" + [guid]::NewGuid().ToString())
New-Item -ItemType Directory -Path $Stage -Force | Out-Null
Expand-Archive -LiteralPath $ZipPath -DestinationPath $Stage -Force
$Payload = Join-Path $Stage "app"
if (-not (Test-Path -LiteralPath $Payload)) {{
    $Payload = $Stage
}}
Copy-Item -Path (Join-Path $Payload "*") -Destination $InstallDir -Recurse -Force
Start-Process -FilePath $Relaunch
"""
    script_path.write_text(payload.strip() + "\n", encoding="utf-8")
    return script_path


def _write_macos_updater_script(base_dir: Path, zip_path: Path, manifest: Dict[str, Any]) -> Path:
    launcher_name = str(manifest.get("launcher_relative_path_macos") or DEFAULT_MANIFEST["launcher_relative_path_macos"])
    relaunch_path = base_dir / launcher_name
    script_path = Path(tempfile.gettempdir()) / "einsatzbericht_apply_update.sh"
    payload = f"""#!/bin/sh
sleep 2
ZIP_PATH={json.dumps(str(zip_path))}
INSTALL_DIR={json.dumps(str(base_dir))}
RELAUNCH={json.dumps(str(relaunch_path))}
STAGE="$(mktemp -d /tmp/einsatzbericht_update.XXXXXX)"
unzip -oq "$ZIP_PATH" -d "$STAGE"
PAYLOAD="$STAGE/app"
if [ ! -d "$PAYLOAD" ]; then
  PAYLOAD="$STAGE"
fi
ditto "$PAYLOAD" "$INSTALL_DIR"
"$RELAUNCH" >/dev/null 2>&1 &
"""
    script_path.write_text(payload, encoding="utf-8")
    script_path.chmod(0o755)
    return script_path


def _launch_external_updater(base_dir: Path, zip_path: Path, manifest: Dict[str, Any]) -> None:
    if sys.platform == "win32":
        script_path = _write_windows_updater_script(base_dir, zip_path, manifest)
        subprocess.Popen(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script_path),
            ],
            creationflags=getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0),
            close_fds=True,
        )
        return

    script_path = _write_macos_updater_script(base_dir, zip_path, manifest)
    subprocess.Popen(["/bin/sh", str(script_path)], close_fds=True)


def maybe_check_for_updates(base_dir: Optional[Path] = None) -> bool:
    if not getattr(sys, "frozen", False):
        return False

    resource_dir = base_dir or _resource_dir()
    install_dir = _install_dir()
    manifest = load_release_manifest(resource_dir)
    runtime = prepare_runtime_environment(resource_dir)
    user_root = runtime["user_root"]

    if not _should_check_for_updates(user_root, manifest):
        return False

    state = _load_update_state(user_root)
    state["last_check_utc"] = dt.datetime.utcnow().isoformat()
    _save_update_state(user_root, state)

    try:
        release = _fetch_latest_release(manifest)
    except Exception:
        return False
    if not release:
        return False

    current_version = str(manifest.get("version") or "").strip()
    latest_version = str(release.get("tag_name") or release.get("name") or "").strip()
    if not latest_version or latest_version == current_version:
        return False

    asset_url = _release_asset_url(release, manifest)
    if not asset_url:
        return False

    body = str(release.get("body") or "").strip()
    changelog = body if body else f"Neue Version: {latest_version}"
    message = textwrap.shorten(
        f"Version {latest_version} ist verfuegbar.\n\nChangelog:\n{changelog}\n\nJetzt aktualisieren?",
        width=3500,
        placeholder="\n...\n",
    )
    if not _ask_yes_no("Update verfuegbar", message):
        return False

    try:
        zip_path = _download_release_asset(asset_url, user_root)
        _launch_external_updater(install_dir, zip_path, manifest)
    except Exception:
        return False
    return True


def app_script_path(base_dir: Optional[Path] = None) -> Path:
    base_dir = base_dir or _resource_dir()
    manifest = load_release_manifest(base_dir)
    main_script = str(manifest.get("main_script") or DEFAULT_MANIFEST["main_script"])
    return (base_dir / main_script).resolve()
