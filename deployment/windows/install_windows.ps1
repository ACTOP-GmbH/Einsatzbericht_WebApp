$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$mainScriptName = "streamlit_einsatzbericht_app_v2_excel_masterdata.py"
$requirementsName = "requirements_einsatzbericht_app_v2_print.txt"
$appSlug = "actop-einsatzbericht-manager"
$installRoot = Join-Path $env:LOCALAPPDATA $appSlug
$installDir = Join-Path $installRoot "app"
$desktopShortcut = Join-Path ([Environment]::GetFolderPath("Desktop")) "Einsatzbericht Manager.lnk"
$startMenuDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs"
$startMenuShortcut = Join-Path $startMenuDir "Einsatzbericht Manager.lnk"
$launcherPath = Join-Path $installDir "run_app.exe"
$launchScriptPath = Join-Path $installDir "launch_app.ps1"
$venvDir = Join-Path $installDir ".venv"
$venvPython = Join-Path $venvDir "Scripts\python.exe"
$powershellPath = Join-Path $PSHOME "powershell.exe"
$pythonCommand = $null
$pythonArgs = @()

function Refresh-ProcessPath {
    $machinePath = [Environment]::GetEnvironmentVariable("Path", "Machine")
    $userPath = [Environment]::GetEnvironmentVariable("Path", "User")
    $paths = @($machinePath, $userPath, $env:Path) | Where-Object { $_ }
    $env:Path = ($paths -join ";")
}

function Show-InfoMessage {
    param(
        [string]$Title,
        [string]$Message
    )

    try {
        Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
        [System.Windows.MessageBox]::Show($Message, $Title) | Out-Null
    } catch {
        Write-Host ""
        Write-Host $Title
        Write-Host $Message
    }
}

function Unblock-Tree {
    param([string]$RootPath)

    if (-not (Test-Path -LiteralPath $RootPath)) {
        return
    }

    try {
        Unblock-File -LiteralPath $RootPath -ErrorAction SilentlyContinue
    } catch {
    }

    Get-ChildItem -LiteralPath $RootPath -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            Unblock-File -LiteralPath $_.FullName -ErrorAction SilentlyContinue
        } catch {
        }
    }
}

function Stop-InstalledApp {
    Get-Process run_app -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

    try {
        Get-CimInstance Win32_Process -Filter "Name = 'python.exe'" -ErrorAction SilentlyContinue |
            Where-Object { $_.CommandLine -and $_.CommandLine -like "*$installDir*" } |
            ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    } catch {
    }

    Start-Sleep -Milliseconds 500
}

function Test-PythonCommand {
    param(
        [string]$Command,
        [string[]]$PrefixArgs
    )

    try {
        $args = @()
        $args += $PrefixArgs
        $args += @("-c", "import sys; print(sys.executable)")
        & $Command @args | Out-Null
        return ($LASTEXITCODE -eq 0)
    } catch {
        return $false
    }
}

function Resolve-PythonCommand {
    Refresh-ProcessPath

    if (Get-Command py -ErrorAction SilentlyContinue) {
        if (Test-PythonCommand -Command "py" -PrefixArgs @("-3")) {
            $script:pythonCommand = "py"
            $script:pythonArgs = @("-3")
            return $true
        }
    }

    $launcherCandidates = @(
        (Join-Path $env:LOCALAPPDATA "Programs\Python\Launcher\py.exe"),
        (Join-Path $env:WINDIR "py.exe")
    )
    foreach ($candidate in $launcherCandidates) {
        if ((Test-Path -LiteralPath $candidate) -and (Test-PythonCommand -Command $candidate -PrefixArgs @("-3"))) {
            $script:pythonCommand = $candidate
            $script:pythonArgs = @("-3")
            return $true
        }
    }

    if (Get-Command python -ErrorAction SilentlyContinue) {
        if (Test-PythonCommand -Command "python" -PrefixArgs @()) {
            $script:pythonCommand = "python"
            $script:pythonArgs = @()
            return $true
        }
    }

    $pythonCandidates = @()
    $pythonRoots = @(
        (Join-Path $env:LOCALAPPDATA "Programs\Python"),
        $env:ProgramFiles,
        ${env:ProgramFiles(x86)}
    ) | Where-Object { $_ }
    foreach ($root in $pythonRoots) {
        $pythonCandidates += Get-ChildItem -Path (Join-Path $root "Python*") -Directory -ErrorAction SilentlyContinue |
            ForEach-Object { Join-Path $_.FullName "python.exe" }
    }

    foreach ($candidate in ($pythonCandidates | Sort-Object -Descending -Unique)) {
        if ($candidate -and (Test-Path -LiteralPath $candidate) -and (Test-PythonCommand -Command $candidate -PrefixArgs @())) {
            $script:pythonCommand = $candidate
            $script:pythonArgs = @()
            return $true
        }
    }

    return $false
}

function Ensure-Python {
    if (Resolve-PythonCommand) {
        return
    }

    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Write-Host "Python wurde nicht gefunden. Versuche Installation ueber winget..."
        winget install --id Python.Python.3.12 -e --scope user --silent --accept-package-agreements --accept-source-agreements
        Refresh-ProcessPath
        Start-Sleep -Seconds 2
        if (Resolve-PythonCommand) {
            return
        }
    }

    throw "Python wurde nicht gefunden und konnte nicht automatisch installiert werden. Bitte Python 3 installieren und install.bat erneut starten."
}

function Invoke-BasePython {
    param([string[]]$ArgsList)

    $args = @()
    $args += $script:pythonArgs
    $args += $ArgsList
    & $script:pythonCommand @args
    if ($LASTEXITCODE -ne 0) {
        throw "Python-Befehl fehlgeschlagen: $($ArgsList -join ' ')"
    }
}

function Install-SourceRuntime {
    $requirementsPath = Join-Path $installDir $requirementsName
    if (-not (Test-Path -LiteralPath $requirementsPath)) {
        throw "Requirements-Datei nicht gefunden: $requirementsPath"
    }

    Ensure-Python
    Invoke-BasePython -ArgsList @("-m", "venv", $venvDir)

    if (-not (Test-Path -LiteralPath $venvPython)) {
        throw "Virtuelle Python-Umgebung konnte nicht erstellt werden: $venvDir"
    }

    & $venvPython -m pip install --upgrade pip
    if ($LASTEXITCODE -ne 0) {
        throw "pip konnte nicht aktualisiert werden."
    }

    & $venvPython -m pip install -r $requirementsPath
    if ($LASTEXITCODE -ne 0) {
        throw "Python-Abhaengigkeiten konnten nicht installiert werden."
    }
}

function Copy-SourcePayload {
    param([string]$SourceDir)

    $files = @(
        $mainScriptName,
        "run_app.py",
        "desktop_runtime.py",
        "release_manifest.json",
        $requirementsName
    )

    foreach ($fileName in $files) {
        $source = Join-Path $SourceDir $fileName
        if (Test-Path -LiteralPath $source) {
            Copy-Item -LiteralPath $source -Destination (Join-Path $installDir $fileName) -Force
        }
    }

    $dataSource = Join-Path $SourceDir "data"
    if (Test-Path -LiteralPath $dataSource) {
        Copy-Item -LiteralPath $dataSource -Destination (Join-Path $installDir "data") -Recurse -Force
    }

    if (-not (Test-Path -LiteralPath (Join-Path $installDir $mainScriptName))) {
        throw "App-Skript nicht gefunden: $mainScriptName"
    }
    if (-not (Test-Path -LiteralPath (Join-Path $installDir "run_app.py"))) {
        throw "run_app.py nicht gefunden. Das Quellpaket ist unvollstaendig."
    }
    if (-not (Test-Path -LiteralPath (Join-Path $installDir "desktop_runtime.py"))) {
        throw "desktop_runtime.py nicht gefunden. Das Quellpaket ist unvollstaendig."
    }
}

function Write-LaunchScript {
    $launchScript = @'
$ErrorActionPreference = "SilentlyContinue"
Add-Type -AssemblyName System.Windows.Forms

$installDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$launcherPath = Join-Path $installDir "run_app.exe"
$venvPython = Join-Path $installDir ".venv\Scripts\python.exe"
$sourceLauncher = Join-Path $installDir "run_app.py"
$stdoutPath = Join-Path $installDir "launch_stdout.log"
$stderrPath = Join-Path $installDir "launch_stderr.log"
$installRoot = Split-Path -Parent $installDir
$portFile = Join-Path $installRoot "logs\server_port.txt"
$port = 8501

$env:STREAMLIT_BROWSER_GATHER_USAGE_STATS = "false"
$env:STREAMLIT_SERVER_HEADLESS = "true"
$env:STREAMLIT_SERVER_SHOW_EMAIL_PROMPT = "false"
$env:STREAMLIT_GLOBAL_DEVELOPMENT_MODE = "false"

$streamlitDir = Join-Path $env:USERPROFILE ".streamlit"
New-Item -ItemType Directory -Path $streamlitDir -Force | Out-Null
Set-Content -LiteralPath (Join-Path $streamlitDir "credentials.toml") -Value "[general]`nemail = `"`"`n" -Encoding UTF8
if (-not (Test-Path -LiteralPath (Join-Path $streamlitDir "config.toml"))) {
    Set-Content -LiteralPath (Join-Path $streamlitDir "config.toml") -Value "[browser]`ngatherUsageStats = false`n`n[server]`nheadless = true`nshowEmailPrompt = false`n" -Encoding UTF8
}

function Test-ServerReady {
    param([int]$Port)
    try {
        $Url = "http://127.0.0.1:$Port/_stcore/health"
        Invoke-WebRequest -UseBasicParsing -Uri $Url -TimeoutSec 2 | Out-Null
        return $true
    } catch {
        return $false
    }
}

function Open-ExistingInstance {
    $ports = @()
    if (Test-Path -LiteralPath $portFile) {
        try {
            $ports += [int](Get-Content -LiteralPath $portFile -Raw)
        } catch {
        }
    }
    $ports += 8501..8520
    foreach ($candidate in ($ports | Select-Object -Unique)) {
        if (Test-ServerReady -Port $candidate) {
            Start-Process "http://127.0.0.1:$candidate" | Out-Null
            return $true
        }
    }
    return $false
}

if (Open-ExistingInstance) {
    exit 0
}

if (Test-Path -LiteralPath $launcherPath) {
    Start-Process -FilePath $launcherPath -WorkingDirectory $installDir -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath | Out-Null
} elseif ((Test-Path -LiteralPath $venvPython) -and (Test-Path -LiteralPath $sourceLauncher)) {
    Start-Process -FilePath $venvPython -ArgumentList "`"$sourceLauncher`"" -WorkingDirectory $installDir -WindowStyle Hidden -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath | Out-Null
} else {
    [System.Windows.Forms.MessageBox]::Show("Die App wurde nicht vollstaendig installiert.", "Startfehler") | Out-Null
    exit 1
}

$deadline = (Get-Date).AddSeconds(45)
while ((Get-Date) -lt $deadline) {
    if (Open-ExistingInstance) {
        exit 0
    }
    Start-Sleep -Milliseconds 500
}

$details = ""
foreach ($path in @($stderrPath, $stdoutPath)) {
    if (Test-Path -LiteralPath $path) {
        $details += (Get-Content -LiteralPath $path -Raw)
    }
}
if ($details) {
    [System.Windows.Forms.MessageBox]::Show("Die App konnte nicht gestartet werden.`n`nDetails:`n$details", "Startfehler") | Out-Null
    exit 1
}

[System.Windows.Forms.MessageBox]::Show("Die App wurde gestartet, aber der Browser konnte die lokale Seite nicht erreichen.`n`nBitte pruefe Firewall/Virenscanner oder starte erneut.", "Startfehler") | Out-Null
'@

    Set-Content -LiteralPath $launchScriptPath -Value $launchScript -Encoding UTF8
    Unblock-File -LiteralPath $launchScriptPath -ErrorAction SilentlyContinue
}

$bundledPayloadDir = Join-Path $scriptRoot "app"
$bundledLauncher = Join-Path $bundledPayloadDir "run_app.exe"
$useBundledPayload = Test-Path -LiteralPath $bundledLauncher

$sourcePayloadCandidates = @(
    $scriptRoot,
    (Split-Path -Parent $scriptRoot),
    (Split-Path -Parent (Split-Path -Parent $scriptRoot)),
    $bundledPayloadDir
)
$sourcePayloadDir = $null
foreach ($candidate in $sourcePayloadCandidates) {
    if ($candidate -and (Test-Path -LiteralPath (Join-Path $candidate $mainScriptName))) {
        $sourcePayloadDir = $candidate
        break
    }
}
if (-not $sourcePayloadDir) {
    $sourcePayloadDir = $scriptRoot
}

New-Item -ItemType Directory -Path $installRoot -Force | Out-Null
New-Item -ItemType Directory -Path $startMenuDir -Force | Out-Null

if (Test-Path -LiteralPath $installDir) {
    Stop-InstalledApp
    Remove-Item -LiteralPath $installDir -Recurse -Force
}
New-Item -ItemType Directory -Path $installDir -Force | Out-Null

if ($useBundledPayload) {
    Unblock-Tree -RootPath $bundledPayloadDir
    Copy-Item -Path (Join-Path $bundledPayloadDir "*") -Destination $installDir -Recurse -Force
    Unblock-Tree -RootPath $installDir
} else {
    Unblock-Tree -RootPath $sourcePayloadDir
    Copy-SourcePayload -SourceDir $sourcePayloadDir
    Unblock-Tree -RootPath $installDir
    Install-SourceRuntime
}

Write-LaunchScript

$wsh = New-Object -ComObject WScript.Shell
foreach ($shortcutPath in @($desktopShortcut, $startMenuShortcut)) {
    $shortcut = $wsh.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $powershellPath
    $shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$launchScriptPath`""
    $shortcut.WorkingDirectory = $installDir
    if (Test-Path -LiteralPath $launcherPath) {
        $shortcut.IconLocation = $launcherPath
    }
    $shortcut.Save()
}

Show-InfoMessage -Title "Installation abgeschlossen" -Message (
    "Die App wurde erfolgreich installiert.`n`n" +
    "Du kannst sie jetzt ueber den Desktop oder das Startmenue starten."
)
