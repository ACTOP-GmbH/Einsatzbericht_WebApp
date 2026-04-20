# Desktop Test Distribution

This project now supports a tester distribution without requiring Git or a local Python installation.

## Packaging model

- The app is bundled with PyInstaller.
- Testers receive a ZIP file.
- The ZIP contains:
  - the bundled app payload in `app/`
  - a platform installer script
  - `README_INSTALL.txt`
- The installer copies the app into the user's profile.
- The seed workbook is copied on first run into a user-local data directory outside the app bundle.
- App updates replace only the bundled code, not the user data.

## Runtime paths

- Windows install root: `%LOCALAPPDATA%\actop-einsatzbericht-manager`
- macOS install root: `~/Library/Application Support/actop-einsatzbericht-manager`
- User workbook: `<install-root>/data/Taetigkeiten_...` is no longer used for active data.
- Active user data is stored in:
  - Windows: `%LOCALAPPDATA%\actop-einsatzbericht-manager\data`
  - macOS: `~/Library/Application Support/actop-einsatzbericht-manager/data`

The runtime entrypoint sets:

- `EINSATZBERICHT_USER_DATA_DIR`
- `EINSATZBERICHT_DEFAULT_EXCEL`

That makes uploads, report imports, and the default workbook resolve into the user profile instead of the repo/app folder.

## Update model

- Updates are checked on app launch.
- The client reads the latest GitHub Release from the repository configured in `release_manifest.json`.
- The release body is shown as the changelog prompt.
- If the user accepts, the release ZIP is downloaded and applied over the installed app folder.
- User data remains untouched because it lives outside the app payload.

## Release workflow

1. Build the desktop payload with PyInstaller.
2. Build the tester ZIP:

```powershell
python scripts/build_desktop_release.py --platform windows
python scripts/build_desktop_release.py --platform macos
```

3. Upload the generated ZIP to a GitHub Release.
4. Put the release notes into the GitHub Release body. That text becomes the user-facing changelog.

## Tester entrypoints

- Windows testers should run `install.bat`.
- `install_windows.ps1` stays as the internal installer implementation, not the user-facing entrypoint.
- macOS testers should run `install_macos.command`.

## Constraint

If the GitHub repository stays private, the release asset still needs to be reachable by testers. The current updater intentionally does not embed Git credentials.
