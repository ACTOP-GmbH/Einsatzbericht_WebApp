from __future__ import annotations

import datetime as dt
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
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
    "runtime_update_check_interval_minutes": 30,
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


def _mark_update_checked(user_root: Path, timestamp_key: str = "last_check_utc", **extra: Any) -> None:
    state = _load_update_state(user_root)
    state[timestamp_key] = dt.datetime.utcnow().isoformat()
    state.update(extra)
    _save_update_state(user_root, state)


def _mark_pending_update_changelog(user_root: Path, version: str, changelog: str) -> None:
    state = _load_update_state(user_root)
    state["pending_changelog_version"] = version
    state["pending_changelog"] = changelog
    state["pending_changelog_created_utc"] = dt.datetime.utcnow().isoformat()
    _save_update_state(user_root, state)


def _should_check_for_updates(
    user_root: Path,
    manifest: Dict[str, Any],
    *,
    timestamp_key: str = "last_check_utc",
    interval_minutes: Optional[int] = None,
) -> bool:
    if not bool(manifest.get("check_updates", True)):
        return False
    state = _load_update_state(user_root)
    if interval_minutes is None:
        interval_minutes = int(float(manifest.get("update_interval_hours", 12) or 12) * 60)
    last_check = str(state.get(timestamp_key) or "").strip()
    if not last_check:
        return True
    try:
        last_dt = dt.datetime.fromisoformat(last_check)
    except Exception:
        return True
    return (dt.datetime.utcnow() - last_dt) >= dt.timedelta(minutes=max(int(interval_minutes), 1))


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


def _show_info(title: str, message: str) -> bool:
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        try:
            messagebox.showinfo(title, message)
            return True
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


def _shorten_dialog_text(message: str, width: int = 3500) -> str:
    if len(message) <= width:
        return message
    return message[: max(0, width - 8)].rstrip() + "\n...\n"


def show_pending_update_changelog(base_dir: Optional[Path] = None) -> bool:
    resource_dir = base_dir or _resource_dir()
    manifest = load_release_manifest(resource_dir)
    current_version = str(manifest.get("version") or "").strip()
    if not current_version or current_version == "dev":
        return False

    user_root = _user_data_root(manifest)
    state = _load_update_state(user_root)
    pending_version = str(state.get("pending_changelog_version") or "").strip()
    if pending_version != current_version:
        return False

    shown_versions = [str(value) for value in (state.get("shown_changelog_versions") or [])]
    if current_version in shown_versions:
        return False

    changelog = str(state.get("pending_changelog") or "").strip()
    if not changelog:
        changelog = f"Version {current_version} wurde installiert."
    message = _shorten_dialog_text(
        f"Version {current_version} wurde installiert.\n\nChangelog:\n{changelog}"
    )
    if not _show_info("Update installiert", message):
        return False

    shown_versions.append(current_version)
    state["shown_changelog_versions"] = sorted(set(shown_versions))
    state.pop("pending_changelog_version", None)
    state.pop("pending_changelog", None)
    state.pop("pending_changelog_created_utc", None)
    _save_update_state(user_root, state)
    return True


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
$ErrorActionPreference = "Stop"
$ProgressPreference = "Continue"
Start-Sleep -Seconds 2
$ZipPath = {json.dumps(str(zip_path))}
$InstallDir = {json.dumps(str(base_dir))}
$Relaunch = {json.dumps(str(relaunch_path))}
$LatestVersion = {json.dumps(str(latest_version))}
$MainScriptName = {json.dumps(main_script)}
$Stage = Join-Path $env:TEMP ("einsatzbericht_update_" + [guid]::NewGuid().ToString())
$script:UpdateActivity = "Einsatzbericht Manager Update"
$script:UpdateStep = 0
$script:UpdateTotalSteps = 7

function Show-UpdateStep {{
    param([string]$Status)

    $script:UpdateStep += 1
    $Percent = [Math]::Min([int](($script:UpdateStep / $script:UpdateTotalSteps) * 100), 99)
    Write-Host ""
    Write-Host "[$($script:UpdateStep)/$script:UpdateTotalSteps] $Status"
    Write-Progress -Activity $script:UpdateActivity -Status $Status -PercentComplete $Percent
}}

function Complete-UpdateProgress {{
    Write-Progress -Activity $script:UpdateActivity -Completed
}}

function Get-InstalledAppProcessIds {{
    $InstallRoot = Split-Path -Parent $InstallDir
    $processIds = New-Object "System.Collections.Generic.HashSet[int]"

    Get-Process -Name "run_app" -ErrorAction SilentlyContinue | ForEach-Object {{
        if ((-not $_.Path) -or $_.Path.StartsWith($InstallDir, [System.StringComparison]::OrdinalIgnoreCase)) {{
            [void]$processIds.Add([int]$_.Id)
        }}
    }}

    foreach ($ProcessName in @("python.exe", "pythonw.exe", "powershell.exe", "pwsh.exe")) {{
        try {{
            Get-CimInstance Win32_Process -Filter "Name = '$ProcessName'" -ErrorAction SilentlyContinue |
                Where-Object {{
                    $cmd = [string]($_.CommandLine)
                    $exe = [string]($_.ExecutablePath)
                    (($cmd -and $cmd.IndexOf($InstallDir, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) -or
                     ($cmd -and $cmd.IndexOf($InstallRoot, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) -or
                     ($exe -and $exe.StartsWith($InstallDir, [System.StringComparison]::OrdinalIgnoreCase)))
                }} |
                ForEach-Object {{ [void]$processIds.Add([int]$_.ProcessId) }}
        }} catch {{
        }}
    }}

    $ids = @()
    foreach ($processId in $processIds) {{
        if ($processId -and $processId -ne $PID) {{
            $ids += [int]$processId
        }}
    }}
    return @($ids | Select-Object -Unique)
}}

function Stop-InstalledApp {{
    $deadline = (Get-Date).AddSeconds(30)

    do {{
        $ids = @(Get-InstalledAppProcessIds)
        if ($ids.Count -eq 0) {{
            return
        }}

        foreach ($processId in $ids) {{
            Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
        }}
        try {{
            Wait-Process -Id $ids -Timeout 3 -ErrorAction SilentlyContinue
        }} catch {{
        }}
        Start-Sleep -Milliseconds 500
    }} while ((Get-Date) -lt $deadline)

    $remaining = @(Get-InstalledAppProcessIds)
    if ($remaining.Count -gt 0) {{
        throw "Die laufende App konnte nicht vollstaendig beendet werden."
    }}
}}

function Test-FileUnlocked {{
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {{
        return $true
    }}

    $stream = $null
    try {{
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        return $true
    }} catch {{
        return $false
    }} finally {{
        if ($stream) {{
            $stream.Close()
        }}
    }}
}}

function Get-LockedAppFiles {{
    $paths = @()
    $launcher = Join-Path $InstallDir "run_app.exe"
    if (Test-Path -LiteralPath $launcher -PathType Leaf) {{
        $paths += $launcher
    }}

    $internalDir = Join-Path $InstallDir "_internal"
    if (Test-Path -LiteralPath $internalDir -PathType Container) {{
        $extensions = @(".dll", ".pyd", ".exe")
        $paths += Get-ChildItem -LiteralPath $internalDir -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object {{ $extensions -contains $_.Extension.ToLowerInvariant() }} |
            ForEach-Object {{ $_.FullName }}
    }}

    $locked = @()
    foreach ($path in ($paths | Sort-Object -Unique)) {{
        if (-not (Test-FileUnlocked -Path $path)) {{
            $locked += $path
        }}
    }}
    return @($locked)
}}

function Wait-InstallDirUnlocked {{
    param([int]$TimeoutSeconds = 45)

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $locked = @()
    do {{
        $locked = @(Get-LockedAppFiles)
        if ($locked.Count -eq 0) {{
            return
        }}
        Stop-InstalledApp
        Start-Sleep -Milliseconds 500
    }} while ((Get-Date) -lt $deadline)

    $sample = ($locked | Select-Object -First 3) -join ", "
    if (-not $sample) {{
        $sample = $InstallDir
    }}
    throw "App-Dateien sind noch durch einen Prozess gesperrt: $sample"
}}

function Copy-PayloadWithRetry {{
    param([string]$PayloadPath)

    for ($Attempt = 1; $Attempt -le 20; $Attempt++) {{
        try {{
            Wait-InstallDirUnlocked -TimeoutSeconds 5
            $PreserveExistingAppData = Test-Path -LiteralPath (Join-Path $InstallDir "data")
            $Items = Get-ChildItem -LiteralPath $PayloadPath -Force -ErrorAction Stop |
                Where-Object {{ -not ($PreserveExistingAppData -and $_.Name -eq "data") }}
            foreach ($Item in $Items) {{
                Copy-Item -LiteralPath $Item.FullName -Destination $InstallDir -Recurse -Force -ErrorAction Stop
            }}
            return
        }} catch {{
            if ($Attempt -eq 20) {{
                throw
            }}
            Stop-InstalledApp
            Start-Sleep -Seconds 1
        }}
    }}
}}

function Test-ServerReady {{
    param([int]$Port)
    try {{
        Invoke-WebRequest -UseBasicParsing -Uri "http://127.0.0.1:$Port/_stcore/health" -TimeoutSec 2 | Out-Null
        return $true
    }} catch {{
        return $false
    }}
}}

function Open-AppWhenReady {{
    $ports = 8501..8510

    foreach ($candidate in ($ports | Select-Object -Unique)) {{
        if (Test-ServerReady -Port $candidate) {{
            Start-Process "http://localhost:$candidate" | Out-Null
            return $true
        }}
    }}
    return $false
}}

function Wait-ForBrowserOpen {{
    param([int]$TimeoutSeconds = 75)

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {{
        if (Open-AppWhenReady) {{
            return $true
        }}
        Start-Sleep -Milliseconds 500
    }}
    return $false
}}

function Start-AppAndWaitForBrowser {{
    param(
        [string]$LaunchScript,
        [string]$Relaunch,
        [string]$VenvPython,
        [string]$SourceLauncher
    )

    Write-Host "Die App wird gestartet. Bitte warten, bis der Browser geoeffnet wurde..."
    if (Test-Path -LiteralPath $LaunchScript) {{
        $process = Start-Process -FilePath "powershell" -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$LaunchScript`"" -Wait -PassThru
        if ($process.ExitCode -ne 0) {{
            throw "Die App konnte nach dem Update nicht gestartet werden."
        }}
        return
    }}

    if (Test-Path -LiteralPath $Relaunch) {{
        Start-Process -FilePath $Relaunch | Out-Null
    }} elseif ((Test-Path -LiteralPath $VenvPython) -and (Test-Path -LiteralPath $SourceLauncher)) {{
        Start-Process -FilePath $VenvPython -ArgumentList "`"$SourceLauncher`"" -WorkingDirectory $InstallDir -WindowStyle Hidden | Out-Null
    }} else {{
        throw "Kein gueltiger App-Starter gefunden."
    }}

    if (-not (Wait-ForBrowserOpen)) {{
        throw "Die App wurde gestartet, aber der Browser konnte nicht automatisch geoeffnet werden."
    }}
}}

try {{
    Show-UpdateStep "Updatepaket wird entpackt"
    New-Item -ItemType Directory -Path $Stage -Force | Out-Null
    Expand-Archive -LiteralPath $ZipPath -DestinationPath $Stage -Force

    Show-UpdateStep "Updateinhalt wird geprueft"
    $Payload = Join-Path $Stage "app"
    if (-not (Test-Path -LiteralPath $Payload)) {{
        $Payload = Get-ChildItem -LiteralPath $Stage -Directory |
            Where-Object {{ Test-Path -LiteralPath (Join-Path $_.FullName $MainScriptName) }} |
            Select-Object -First 1 -ExpandProperty FullName
    }}
    if (-not $Payload) {{
        $Payload = $Stage
    }}

    Show-UpdateStep "Laufende App wird beendet"
    Stop-InstalledApp
    Wait-InstallDirUnlocked -TimeoutSeconds 45

    Show-UpdateStep "Neue App-Dateien werden kopiert"
    Copy-PayloadWithRetry -PayloadPath $Payload

    Show-UpdateStep "Versionsinformationen werden aktualisiert"
    $ManifestPath = Join-Path $InstallDir "release_manifest.json"
    if (Test-Path -LiteralPath $ManifestPath) {{
        try {{
            $Manifest = Get-Content -LiteralPath $ManifestPath -Raw | ConvertFrom-Json
            $Manifest.version = $LatestVersion
            $Manifest | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $ManifestPath -Encoding UTF8
        }} catch {{
        }}
    }}

    Show-UpdateStep "Abhaengigkeiten werden geprueft"
    $RequirementsPath = Join-Path $InstallDir "requirements_einsatzbericht_app_v2_print.txt"
    $VenvPython = Join-Path $InstallDir ".venv\\Scripts\\python.exe"
    if ((Test-Path -LiteralPath $VenvPython) -and (Test-Path -LiteralPath $RequirementsPath)) {{
        & $VenvPython -m pip install -r $RequirementsPath
        if ($LASTEXITCODE -ne 0) {{
            throw "Python-Abhaengigkeiten konnten nicht aktualisiert werden."
        }}
    }}

    Show-UpdateStep "App wird gestartet"
    $LaunchScript = Join-Path $InstallDir "launch_app.ps1"
    $SourceLauncher = Join-Path $InstallDir "run_app.py"
    Complete-UpdateProgress
    Write-Host ""
    Write-Host "Update abgeschlossen. Die App wird jetzt gestartet."
    Start-AppAndWaitForBrowser -LaunchScript $LaunchScript -Relaunch $Relaunch -VenvPython $VenvPython -SourceLauncher $SourceLauncher
    Write-Host "App ist gestartet. Der Browser wurde geoeffnet."
    Start-Sleep -Seconds 2
}} catch {{
    Complete-UpdateProgress
    Write-Host ""
    Write-Host "Update fehlgeschlagen:"
    Write-Host $_.Exception.Message
    Write-Host ""
    Read-Host "Enter druecken, um dieses Fenster zu schliessen"
    throw
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
set -eu

notify() {{
  osascript -e "display notification \\"$1\\" with title \\"Einsatzbericht Manager\\"" >/dev/null 2>&1 || true
}}

fail() {{
  echo ""
  echo "Update fehlgeschlagen."
  notify "Update fehlgeschlagen."
  echo "Dieses Fenster kann geschlossen werden."
  exit 1
}}

trap fail ERR

echo "Einsatzbericht Manager Update"
echo "Bitte dieses Fenster nicht schliessen."
notify "Update wird installiert."
sleep 2
ZIP_PATH={json.dumps(str(zip_path))}
INSTALL_DIR={json.dumps(str(base_dir))}
RELAUNCH={json.dumps(str(relaunch_path))}
LATEST_VERSION={json.dumps(str(latest_version))}
MAIN_SCRIPT_NAME={json.dumps(main_script)}
STAGE="$(mktemp -d /tmp/einsatzbericht_update.XXXXXX)"
echo ""
echo "[1/6] Updatepaket wird entpackt"
unzip -oq "$ZIP_PATH" -d "$STAGE"
echo "[2/6] Updateinhalt wird geprueft"
PAYLOAD="$STAGE/app"
if [ ! -d "$PAYLOAD" ]; then
  FOUND="$(find "$STAGE" -maxdepth 2 -type f -name "$MAIN_SCRIPT_NAME" -print -quit)"
  if [ -n "$FOUND" ]; then
    PAYLOAD="$(dirname "$FOUND")"
  else
    PAYLOAD="$STAGE"
  fi
fi

stop_installed_app() {{
  echo "[3/6] Laufende App wird beendet"
  if [ -n "$RELAUNCH" ]; then
    pkill -f "$RELAUNCH" >/dev/null 2>&1 || true
  fi
  pkill -f "$INSTALL_DIR/run_app.py" >/dev/null 2>&1 || true
  pkill -f "$INSTALL_DIR/.venv" >/dev/null 2>&1 || true
}}

copy_payload_with_retry() {{
  attempt=1
  while [ "$attempt" -le 12 ]; do
    if [ -d "$INSTALL_DIR/data" ]; then
      ok=1
      for item in "$PAYLOAD"/* "$PAYLOAD"/.[!.]* "$PAYLOAD"/..?*; do
        [ -e "$item" ] || continue
        name="$(basename "$item")"
        [ "$name" = "data" ] && continue
        if ! ditto "$item" "$INSTALL_DIR/$name"; then
          ok=0
          break
        fi
      done
      [ "$ok" -eq 1 ] && return 0
    elif ditto "$PAYLOAD" "$INSTALL_DIR"; then
      return 0
    fi
    stop_installed_app
    sleep 1
    attempt=$((attempt + 1))
  done
  return 1
}}

server_ready() {{
  port="$1"
  curl -fsS --max-time 2 "http://localhost:$port/_stcore/health" >/dev/null 2>&1
}}

open_app_when_ready() {{
  install_root="$(dirname "$INSTALL_DIR")"
  port_file="$install_root/logs/server_port.txt"
  ports=""
  if [ -f "$port_file" ]; then
    ports="$(cat "$port_file" 2>/dev/null || true)"
  fi
  ports="$ports 8501 8502 8503 8504 8505 8506 8507 8508 8509 8510 8511 8512 8513 8514 8515 8516 8517 8518 8519 8520"
  for port in $ports; do
    case "$port" in
      ''|*[!0-9]*) continue ;;
    esac
    if server_ready "$port"; then
      open "http://localhost:$port" >/dev/null 2>&1 || true
      return 0
    fi
  done
  return 1
}}

wait_for_browser_open() {{
  deadline=$(( $(date +%s) + 75 ))
  while [ "$(date +%s)" -lt "$deadline" ]; do
    if open_app_when_ready; then
      return 0
    fi
    sleep 1
  done
  return 1
}}

stop_installed_app
echo "[4/6] Neue App-Dateien werden kopiert"
copy_payload_with_retry
echo "[5/6] Versionsinformationen werden aktualisiert"
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
echo "[6/6] App wird gestartet"
echo ""
echo "Update abgeschlossen. Die App wird jetzt gestartet."
notify "Update abgeschlossen. Die App wird gestartet."
if [ -x "$RELAUNCH" ]; then
  "$RELAUNCH" >/dev/null 2>&1 &
else
  python3 "$INSTALL_DIR/run_app.py" >/dev/null 2>&1 &
fi
echo "Bitte warten, bis der Browser geoeffnet wurde..."
if wait_for_browser_open; then
  echo "App ist gestartet. Der Browser wurde geoeffnet."
  notify "App ist gestartet."
else
  echo "Die App wurde gestartet, aber der Browser konnte nicht automatisch geoeffnet werden."
  echo "Bitte starte die App erneut oder oeffne localhost manuell."
  notify "App gestartet, Browser konnte nicht automatisch geoeffnet werden."
fi
sleep 3
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
            creationflags=getattr(subprocess, "CREATE_NEW_CONSOLE", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0),
            close_fds=True,
        )
        return

    script_path = _write_macos_updater_script(base_dir, zip_path, manifest, latest_version)
    try:
        subprocess.Popen(["open", "-a", "Terminal", str(script_path)], close_fds=True)
    except Exception:
        subprocess.Popen(["/bin/sh", str(script_path)], close_fds=True)


def _runtime_update_interval_minutes(manifest: Dict[str, Any]) -> int:
    try:
        return max(int(manifest.get("runtime_update_check_interval_minutes", 30) or 30), 1)
    except Exception:
        return 30


def _available_update_payload(
    manifest: Dict[str, Any],
    current_version: str,
) -> Optional[Dict[str, str]]:
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

    if not update_url and _is_commit_version(current_version) and not getattr(sys, "frozen", False):
        try:
            latest_commit = _fetch_latest_commit(manifest)
        except Exception:
            latest_commit = None
        if latest_commit:
            latest_sha = str(latest_commit["sha"])
            if latest_sha.startswith(current_version):
                return None
            archive_url = _source_archive_url(manifest, latest_sha)
            if archive_url:
                update_url = archive_url
                latest_version = str(latest_commit["short_sha"])
                download_name = f"EinsatzberichtManager-source-{latest_version}.zip"
                changelog = str(latest_commit.get("message") or "").strip() or f"Neuer Commit: {latest_version}"

    if not update_url or not latest_version:
        return None

    return {
        "update_url": update_url,
        "latest_version": latest_version,
        "changelog": changelog,
        "download_name": download_name,
        "current_version": current_version,
    }


def check_for_update_info(base_dir: Optional[Path] = None, *, force: bool = False) -> Dict[str, Any]:
    resource_dir = base_dir or _resource_dir()
    manifest = load_release_manifest(resource_dir)
    current_version = str(manifest.get("version") or "").strip()
    if not getattr(sys, "frozen", False) and (not current_version or current_version == "dev"):
        return {"available": False, "reason": "dev"}

    runtime = prepare_runtime_environment(resource_dir)
    user_root = runtime["user_root"]
    interval_minutes = _runtime_update_interval_minutes(manifest)
    timestamp_key = "last_runtime_check_utc"

    if not force and not _should_check_for_updates(
        user_root,
        manifest,
        timestamp_key=timestamp_key,
        interval_minutes=interval_minutes,
    ):
        return {"available": False, "reason": "throttled"}

    try:
        payload = _available_update_payload(manifest, current_version)
    except Exception as exc:
        _mark_update_checked(
            user_root,
            timestamp_key=timestamp_key,
            last_runtime_update_result="failed",
            last_runtime_update_error=str(exc),
        )
        return {"available": False, "reason": "failed", "error": str(exc)}

    if not payload:
        _mark_update_checked(user_root, timestamp_key=timestamp_key, last_runtime_update_result="no_update")
        return {"available": False, "reason": "no_update"}

    _mark_update_checked(
        user_root,
        timestamp_key=timestamp_key,
        last_runtime_update_result="available",
        latest_version=payload["latest_version"],
    )
    return {"available": True, **payload}


def start_update_from_info(update_info: Dict[str, Any], base_dir: Optional[Path] = None) -> tuple[bool, str]:
    resource_dir = base_dir or _resource_dir()
    install_dir = _install_dir()
    manifest = load_release_manifest(resource_dir)
    runtime = prepare_runtime_environment(resource_dir)
    user_root = runtime["user_root"]

    update_url = str(update_info.get("update_url") or "").strip()
    latest_version = str(update_info.get("latest_version") or "").strip()
    changelog = str(update_info.get("changelog") or "").strip()
    download_name = str(update_info.get("download_name") or "update.zip").strip() or "update.zip"
    if not update_url or not latest_version:
        return False, "Update-Information ist unvollstaendig."

    try:
        zip_path = _download_update_zip(update_url, user_root, download_name)
        _mark_pending_update_changelog(user_root, latest_version, changelog)
        _launch_external_updater(install_dir, zip_path, manifest, latest_version)
        _mark_update_checked(user_root, last_update_result="started", latest_version=latest_version)
    except Exception as exc:
        _mark_update_checked(user_root, last_update_result="failed", last_update_error=str(exc))
        return False, str(exc)
    return True, "Update wird installiert. Die App startet danach neu."


def maybe_check_for_updates(base_dir: Optional[Path] = None, *, force: bool = False) -> bool:
    resource_dir = base_dir or _resource_dir()
    manifest = load_release_manifest(resource_dir)
    current_version = str(manifest.get("version") or "").strip()
    if not getattr(sys, "frozen", False) and (not current_version or current_version == "dev"):
        return False
    runtime = prepare_runtime_environment(resource_dir)
    user_root = runtime["user_root"]

    if not force and not _should_check_for_updates(user_root, manifest):
        return False

    update_info = _available_update_payload(manifest, current_version)
    if not update_info:
        _mark_update_checked(user_root, last_update_result="no_update")
        return False

    latest_version = update_info["latest_version"]
    changelog = update_info["changelog"]
    message = _shorten_dialog_text(
        f"Version {latest_version} ist verfuegbar.\n\nChangelog:\n{changelog}\n\nJetzt aktualisieren?",
    )
    if not _ask_yes_no("Update verfuegbar", message):
        _mark_update_checked(user_root, last_update_result="declined", latest_version=latest_version)
        return False

    ok, _message = start_update_from_info(update_info, resource_dir)
    if not ok:
        return False
    return True


def app_script_path(base_dir: Optional[Path] = None) -> Path:
    base_dir = base_dir or _resource_dir()
    manifest = load_release_manifest(base_dir)
    main_script = str(manifest.get("main_script") or DEFAULT_MANIFEST["main_script"])
    return (base_dir / main_script).resolve()
