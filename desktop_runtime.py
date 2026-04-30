from __future__ import annotations

import datetime as dt
import json
import os
import re
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
    "github_branch": "master",
    "launcher_relative_path_windows": "run_app.exe",
    "launcher_relative_path_macos": "run_app",
    "check_updates": True,
    "update_interval_hours": 12,
}


def configure_streamlit_runtime() -> None:
    """Disable Streamlit first-run prompts and telemetry for desktop launches."""
    os.environ.setdefault("STREAMLIT_BROWSER_GATHER_USAGE_STATS", "false")
    os.environ.setdefault("STREAMLIT_SERVER_HEADLESS", "true")
    os.environ.setdefault("STREAMLIT_SERVER_SHOW_EMAIL_PROMPT", "false")
    os.environ.setdefault("STREAMLIT_GLOBAL_DEVELOPMENT_MODE", "false")

    try:
        streamlit_dir = Path.home() / ".streamlit"
        streamlit_dir.mkdir(parents=True, exist_ok=True)

        credentials_path = streamlit_dir / "credentials.toml"
        credentials_path.write_text('[general]\nemail = ""\n', encoding="utf-8")

        config_path = streamlit_dir / "config.toml"
        if not config_path.exists():
            config_path.write_text(
                "[browser]\ngatherUsageStats = false\n\n"
                "[server]\nheadless = true\nshowEmailPrompt = false\n",
                encoding="utf-8",
            )
    except Exception:
        pass


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
    if seed_source.exists() and (not seed_target.exists() or seed_target.stat().st_size == 0):
        temp_target = seed_target.with_suffix(seed_target.suffix + ".tmp")
        try:
            shutil.copy2(seed_source, temp_target)
            temp_target.replace(seed_target)
        finally:
            try:
                if temp_target.exists():
                    temp_target.unlink()
            except Exception:
                pass

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


def _mark_update_checked(user_root: Path, **extra: Any) -> None:
    state = _load_update_state(user_root)
    state["last_check_utc"] = dt.datetime.utcnow().isoformat()
    state.update(extra)
    _save_update_state(user_root, state)


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


def _fetch_latest_commit(manifest: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    repo = str(manifest.get("github_repo") or "").strip()
    branch = str(manifest.get("github_branch") or DEFAULT_MANIFEST["github_branch"]).strip()
    if not repo or not branch:
        return None
    url = f"https://api.github.com/repos/{repo}/commits/{branch}"
    req = urllib.request.Request(url, headers={"User-Agent": "einsatzbericht-manager-updater"})
    with urllib.request.urlopen(req, timeout=10) as response:
        data = json.loads(response.read().decode("utf-8"))
    sha = str(data.get("sha") or "").strip()
    if not sha:
        return None
    commit = data.get("commit") or {}
    return {
        "sha": sha,
        "short_sha": sha[:7],
        "message": str(commit.get("message") or "").strip(),
        "html_url": str(data.get("html_url") or "").strip(),
    }


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


def _source_archive_url(manifest: Dict[str, Any], commit_sha: str) -> Optional[str]:
    repo = str(manifest.get("github_repo") or "").strip()
    if not repo or not commit_sha:
        return None
    return f"https://api.github.com/repos/{repo}/zipball/{commit_sha}"


def _is_commit_version(value: str) -> bool:
    return bool(re.fullmatch(r"[0-9a-fA-F]{7,40}", value.strip()))


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


def _download_update_zip(url: str, user_root: Path, file_name: str = "update.zip") -> Path:
    downloads_dir = user_root / "downloads"
    downloads_dir.mkdir(parents=True, exist_ok=True)
    file_name = Path(url.split("?")[0]).name or file_name
    if not file_name.lower().endswith(".zip"):
        file_name = f"{file_name}.zip"
    target = downloads_dir / file_name
    req = urllib.request.Request(url, headers={"User-Agent": "einsatzbericht-manager-updater"})
    with urllib.request.urlopen(req, timeout=30) as response:
        target.write_bytes(response.read())
    return target


def _write_windows_updater_script(base_dir: Path, zip_path: Path, manifest: Dict[str, Any], latest_version: str) -> Path:
    launcher_name = str(manifest.get("launcher_relative_path_windows") or DEFAULT_MANIFEST["launcher_relative_path_windows"])
    relaunch_path = base_dir / launcher_name
    main_script = str(manifest.get("main_script") or DEFAULT_MANIFEST["main_script"])
    script_path = Path(tempfile.gettempdir()) / "einsatzbericht_apply_update.ps1"
    payload = f"""
Start-Sleep -Seconds 2
$ZipPath = {json.dumps(str(zip_path))}
$InstallDir = {json.dumps(str(base_dir))}
$Relaunch = {json.dumps(str(relaunch_path))}
$LatestVersion = {json.dumps(str(latest_version))}
$MainScriptName = {json.dumps(main_script)}
$Stage = Join-Path $env:TEMP ("einsatzbericht_update_" + [guid]::NewGuid().ToString())
New-Item -ItemType Directory -Path $Stage -Force | Out-Null
Expand-Archive -LiteralPath $ZipPath -DestinationPath $Stage -Force
$Payload = Join-Path $Stage "app"
if (-not (Test-Path -LiteralPath $Payload)) {{
    $Payload = Get-ChildItem -LiteralPath $Stage -Directory |
        Where-Object {{ Test-Path -LiteralPath (Join-Path $_.FullName $MainScriptName) }} |
        Select-Object -First 1 -ExpandProperty FullName
}}
if (-not $Payload) {{
    $Payload = $Stage
}}
Copy-Item -Path (Join-Path $Payload "*") -Destination $InstallDir -Recurse -Force
$ManifestPath = Join-Path $InstallDir "release_manifest.json"
if (Test-Path -LiteralPath $ManifestPath) {{
    try {{
        $Manifest = Get-Content -LiteralPath $ManifestPath -Raw | ConvertFrom-Json
        $Manifest.version = $LatestVersion
        $Manifest | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $ManifestPath -Encoding UTF8
    }} catch {{
    }}
}}
$RequirementsPath = Join-Path $InstallDir "requirements_einsatzbericht_app_v2_print.txt"
$VenvPython = Join-Path $InstallDir ".venv\\Scripts\\python.exe"
if ((Test-Path -LiteralPath $VenvPython) -and (Test-Path -LiteralPath $RequirementsPath)) {{
    Start-Process -FilePath $VenvPython -ArgumentList "-m", "pip", "install", "-r", "`"$RequirementsPath`"" -Wait -WindowStyle Hidden
}}
$LaunchScript = Join-Path $InstallDir "launch_app.ps1"
$SourceLauncher = Join-Path $InstallDir "run_app.py"
if (Test-Path -LiteralPath $Relaunch) {{
    Start-Process -FilePath $Relaunch
}} elseif (Test-Path -LiteralPath $LaunchScript) {{
    Start-Process -FilePath "powershell" -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden", "-File", "`"$LaunchScript`""
}} elseif ((Test-Path -LiteralPath $VenvPython) -and (Test-Path -LiteralPath $SourceLauncher)) {{
    Start-Process -FilePath $VenvPython -ArgumentList "`"$SourceLauncher`"" -WorkingDirectory $InstallDir -WindowStyle Hidden
}}
"""
    script_path.write_text(payload.strip() + "\n", encoding="utf-8")
    return script_path


def _write_macos_updater_script(base_dir: Path, zip_path: Path, manifest: Dict[str, Any], latest_version: str) -> Path:
    launcher_name = str(manifest.get("launcher_relative_path_macos") or DEFAULT_MANIFEST["launcher_relative_path_macos"])
    relaunch_path = base_dir / launcher_name
    main_script = str(manifest.get("main_script") or DEFAULT_MANIFEST["main_script"])
    script_path = Path(tempfile.gettempdir()) / "einsatzbericht_apply_update.sh"
    payload = f"""#!/bin/sh
sleep 2
ZIP_PATH={json.dumps(str(zip_path))}
INSTALL_DIR={json.dumps(str(base_dir))}
RELAUNCH={json.dumps(str(relaunch_path))}
LATEST_VERSION={json.dumps(str(latest_version))}
MAIN_SCRIPT_NAME={json.dumps(main_script)}
STAGE="$(mktemp -d /tmp/einsatzbericht_update.XXXXXX)"
unzip -oq "$ZIP_PATH" -d "$STAGE"
PAYLOAD="$STAGE/app"
if [ ! -d "$PAYLOAD" ]; then
  FOUND="$(find "$STAGE" -maxdepth 2 -type f -name "$MAIN_SCRIPT_NAME" -print -quit)"
  if [ -n "$FOUND" ]; then
    PAYLOAD="$(dirname "$FOUND")"
  else
    PAYLOAD="$STAGE"
  fi
fi
ditto "$PAYLOAD" "$INSTALL_DIR"
python3 - <<PY >/dev/null 2>&1
import json
from pathlib import Path
p = Path("$INSTALL_DIR") / "release_manifest.json"
try:
    data = json.loads(p.read_text(encoding="utf-8"))
    data["version"] = "$LATEST_VERSION"
    p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
except Exception:
    pass
PY
if [ -x "$RELAUNCH" ]; then
  "$RELAUNCH" >/dev/null 2>&1 &
else
  python3 "$INSTALL_DIR/run_app.py" >/dev/null 2>&1 &
fi
"""
    script_path.write_text(payload, encoding="utf-8")
    script_path.chmod(0o755)
    return script_path


def _launch_external_updater(base_dir: Path, zip_path: Path, manifest: Dict[str, Any], latest_version: str) -> None:
    if sys.platform == "win32":
        script_path = _write_windows_updater_script(base_dir, zip_path, manifest, latest_version)
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

    script_path = _write_macos_updater_script(base_dir, zip_path, manifest, latest_version)
    subprocess.Popen(["/bin/sh", str(script_path)], close_fds=True)


def maybe_check_for_updates(base_dir: Optional[Path] = None) -> bool:
    resource_dir = base_dir or _resource_dir()
    install_dir = _install_dir()
    manifest = load_release_manifest(resource_dir)
    current_version = str(manifest.get("version") or "").strip()
    if not getattr(sys, "frozen", False) and (not current_version or current_version == "dev"):
        return False
    runtime = prepare_runtime_environment(resource_dir)
    user_root = runtime["user_root"]

    if not _should_check_for_updates(user_root, manifest):
        return False

    update_url: Optional[str] = None
    latest_version = ""
    changelog = ""
    download_name = "update.zip"
    try:
        release = _fetch_latest_release(manifest)
    except Exception:
        release = None
    if release:
        release_version = str(release.get("tag_name") or release.get("name") or "").strip()
        asset_url = _release_asset_url(release, manifest)
        if release_version and release_version != current_version and asset_url:
            update_url = asset_url
            latest_version = release_version
            body = str(release.get("body") or "").strip()
            changelog = body if body else f"Neue Version: {latest_version}"

    if not update_url and _is_commit_version(current_version):
        try:
            latest_commit = _fetch_latest_commit(manifest)
        except Exception:
            latest_commit = None
        if latest_commit:
            latest_sha = str(latest_commit["sha"])
            if latest_sha.startswith(current_version):
                _mark_update_checked(user_root, last_update_result="current", latest_version=latest_commit["short_sha"])
                return False
            archive_url = _source_archive_url(manifest, latest_sha)
            if archive_url:
                update_url = archive_url
                latest_version = str(latest_commit["short_sha"])
                download_name = f"EinsatzberichtManager-source-{latest_version}.zip"
                changelog = str(latest_commit.get("message") or "").strip() or f"Neuer Commit: {latest_version}"

    if not update_url or not latest_version:
        _mark_update_checked(user_root, last_update_result="no_update")
        return False

    message = textwrap.shorten(
        f"Version {latest_version} ist verfuegbar.\n\nChangelog:\n{changelog}\n\nJetzt aktualisieren?",
        width=3500,
        placeholder="\n...\n",
    )
    if not _ask_yes_no("Update verfuegbar", message):
        _mark_update_checked(user_root, last_update_result="declined", latest_version=latest_version)
        return False

    try:
        zip_path = _download_update_zip(update_url, user_root, download_name)
        _launch_external_updater(install_dir, zip_path, manifest, latest_version)
        _mark_update_checked(user_root, last_update_result="started", latest_version=latest_version)
    except Exception as exc:
        _mark_update_checked(user_root, last_update_result="failed", last_update_error=str(exc))
        return False
    return True


def app_script_path(base_dir: Optional[Path] = None) -> Path:
    base_dir = base_dir or _resource_dir()
    manifest = load_release_manifest(base_dir)
    main_script = str(manifest.get("main_script") or DEFAULT_MANIFEST["main_script"])
    return (base_dir / main_script).resolve()
