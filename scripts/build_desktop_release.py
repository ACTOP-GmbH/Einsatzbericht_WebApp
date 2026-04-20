from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
DIST_DIR = ROOT / "dist" / "run_app"
MANIFEST_PATH = ROOT / "release_manifest.json"


def _seed_workbook_path() -> Path:
    for name in ["Tätigkeiten_Überblick.xlsx", "Taetigkeiten_Ueberblick.xlsx"]:
        candidate = ROOT / "data" / name
        if candidate.exists():
            return candidate
    return ROOT / "data" / "Taetigkeiten_Ueberblick.xlsx"


def _git_version() -> str:
    try:
        return subprocess.check_output(
            ["git", "rev-parse", "--short", "HEAD"],
            cwd=ROOT,
            text=True,
        ).strip()
    except Exception:
        return "dev"


def _load_manifest() -> dict:
    data = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
    return data


def _platform_asset_name(platform_name: str) -> str:
    return "EinsatzberichtManager-windows.zip" if platform_name == "windows" else "EinsatzberichtManager-macos.zip"


def _installer_name(platform_name: str) -> str:
    return "install_windows.ps1" if platform_name == "windows" else "install_macos.command"


def _installer_source(platform_name: str) -> Path:
    if platform_name == "windows":
        return ROOT / "deployment" / "windows" / "install_windows.ps1"
    return ROOT / "deployment" / "macos" / "install_macos.command"


def _windows_wrapper_source() -> Path:
    return ROOT / "deployment" / "windows" / "install.bat"


def _copy_runtime_compat_files(app_dir: Path, seed_workbook: Path) -> None:
    internal_dir = app_dir / "_internal"
    if not internal_dir.exists():
        return

    for relative_path in [
        Path("release_manifest.json"),
        Path("streamlit_einsatzbericht_app_v2_excel_masterdata.py"),
    ]:
        source = internal_dir / relative_path
        if source.exists():
            target = app_dir / relative_path
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, target)

    target_data_dir = app_dir / "data"
    target_data_dir.mkdir(parents=True, exist_ok=True)
    internal_seed = internal_dir / "data" / seed_workbook.name
    if internal_seed.exists():
        shutil.copy2(internal_seed, target_data_dir / seed_workbook.name)


def build_release(platform_name: str, version: str, payload_dir: Path, output_dir: Path) -> Path:
    seed_workbook = _seed_workbook_path()
    if not payload_dir.exists():
        raise FileNotFoundError(f"Build payload not found: {payload_dir}")
    if not seed_workbook.exists():
        raise FileNotFoundError(f"Seed workbook not found: {seed_workbook}")

    staging_dir = output_dir / f"release_{platform_name}"
    app_dir = staging_dir / "app"
    if staging_dir.exists():
        shutil.rmtree(staging_dir)
    app_dir.mkdir(parents=True, exist_ok=True)

    shutil.copytree(payload_dir, app_dir, dirs_exist_ok=True)
    (app_dir / "data").mkdir(parents=True, exist_ok=True)
    shutil.copy2(seed_workbook, app_dir / "data" / seed_workbook.name)
    _copy_runtime_compat_files(app_dir, seed_workbook)

    manifest = _load_manifest()
    manifest["version"] = version
    manifest["release_asset_windows"] = _platform_asset_name("windows")
    manifest["release_asset_macos"] = _platform_asset_name("macos")
    (app_dir / "release_manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    installer_source = _installer_source(platform_name)
    shutil.copy2(installer_source, staging_dir / _installer_name(platform_name))
    if platform_name == "windows":
        shutil.copy2(_windows_wrapper_source(), staging_dir / "install.bat")

    readme = staging_dir / "README_INSTALL.txt"
    starter_name = "install.bat" if platform_name == "windows" else _installer_name(platform_name)
    readme.write_text(
        "\n".join(
            [
                "Einsatzbericht Manager Testdistribution",
                "",
                f"Version: {version}",
                "",
                "1. Entpacke diese ZIP-Datei.",
                f"2. Starte {starter_name}.",
                "3. Die App kopiert sich in dein Benutzerprofil und legt eine Startverknuepfung an.",
                "4. Updates werden beim Start ueber GitHub Releases geprueft.",
                "",
                "Hinweis:",
                "Wenn das Repository privat bleibt, muessen die Release-ZIPs oeffentlich oder ueber einen anderen Download-Kanal erreichbar sein.",
            ]
        ),
        encoding="utf-8",
    )

    zip_path = output_dir / _platform_asset_name(platform_name)
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in staging_dir.rglob("*"):
            if path.is_file():
                zf.write(path, path.relative_to(staging_dir))
    return zip_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Build a tester ZIP for the desktop app.")
    parser.add_argument("--platform", choices=["windows", "macos"], required=True)
    parser.add_argument("--version", default=_git_version())
    parser.add_argument("--payload-dir", default=str(DIST_DIR))
    parser.add_argument("--output-dir", default=str(ROOT / "release"))
    args = parser.parse_args()

    zip_path = build_release(
        platform_name=args.platform,
        version=args.version,
        payload_dir=Path(args.payload_dir).resolve(),
        output_dir=Path(args.output_dir).resolve(),
    )
    print(f"Created release ZIP: {zip_path}")


if __name__ == "__main__":
    main()
