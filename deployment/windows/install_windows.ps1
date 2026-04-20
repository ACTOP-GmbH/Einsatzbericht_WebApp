$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$packageAppDir = Join-Path $scriptRoot "app"
if (-not (Test-Path -LiteralPath $packageAppDir)) {
    throw "Package payload not found: $packageAppDir"
}

$appSlug = "actop-einsatzbericht-manager"
$installRoot = Join-Path $env:LOCALAPPDATA $appSlug
$installDir = Join-Path $installRoot "app"
$desktopShortcut = Join-Path ([Environment]::GetFolderPath("Desktop")) "Einsatzbericht Manager.lnk"
$startMenuDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs"
$startMenuShortcut = Join-Path $startMenuDir "Einsatzbericht Manager.lnk"
$launcherPath = Join-Path $installDir "run_app.exe"

New-Item -ItemType Directory -Path $installRoot -Force | Out-Null

if (Test-Path -LiteralPath $installDir) {
    Remove-Item -LiteralPath $installDir -Recurse -Force
}

Copy-Item -LiteralPath $packageAppDir -Destination $installDir -Recurse -Force

if (-not (Test-Path -LiteralPath $launcherPath)) {
    throw "Launcher not found after install: $launcherPath"
}

$wsh = New-Object -ComObject WScript.Shell
foreach ($shortcutPath in @($desktopShortcut, $startMenuShortcut)) {
    $shortcut = $wsh.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $launcherPath
    $shortcut.WorkingDirectory = $installDir
    $shortcut.IconLocation = $launcherPath
    $shortcut.Save()
}

Start-Process -FilePath $launcherPath

