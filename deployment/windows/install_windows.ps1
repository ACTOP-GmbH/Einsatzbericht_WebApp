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
$launchWrapperPath = Join-Path $installDir "launch_app.vbs"
$venvDir = Join-Path $installDir ".venv"
$venvPython = Join-Path $venvDir "Scripts\python.exe"
$powershellPath = Join-Path $PSHOME "powershell.exe"
$wscriptPath = Join-Path $env:WINDIR "System32\wscript.exe"
$pythonCommand = $null
$pythonArgs = @()
$script:installActivity = "Einsatzbericht Manager Installation"
$script:installStep = 0
$script:installTotalSteps = 7
$script:preserveExistingAppData = $false

function Show-InstallStep {
    param([string]$Status)

    $script:installStep += 1
    $percent = [Math]::Min([int](($script:installStep / $script:installTotalSteps) * 100), 99)
    Write-Host ""
    Write-Host "[$script:installStep/$script:installTotalSteps] $Status"
    Write-Progress -Activity $script:installActivity -Status $Status -PercentComplete $percent
}

function Show-InstallStatus {
    param([string]$Status)

    $percent = [Math]::Min([int](($script:installStep / $script:installTotalSteps) * 100), 99)
    Write-Host "  $Status"
    Write-Progress -Activity $script:installActivity -Status $Status -PercentComplete $percent
}

function Complete-InstallProgress {
    Write-Progress -Activity $script:installActivity -Completed
}

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

function Get-InstalledAppProcessIds {
    $ids = New-Object "System.Collections.Generic.HashSet[int]"

    Get-Process run_app -ErrorAction SilentlyContinue | ForEach-Object {
        if ((-not $_.Path) -or $_.Path.StartsWith($installDir, [System.StringComparison]::OrdinalIgnoreCase)) {
            [void]$ids.Add([int]$_.Id)
        }
    }

    foreach ($processName in @("python.exe", "pythonw.exe", "powershell.exe", "pwsh.exe")) {
        try {
            Get-CimInstance Win32_Process -Filter "Name = '$processName'" -ErrorAction SilentlyContinue |
                Where-Object {
                    $cmd = [string]($_.CommandLine)
                    $exe = [string]($_.ExecutablePath)
                    ($cmd.IndexOf($installDir, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) -or
                    ($cmd.IndexOf($installRoot, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) -or
                    ($exe -and $exe.StartsWith($installDir, [System.StringComparison]::OrdinalIgnoreCase))
                } |
                ForEach-Object { [void]$ids.Add([int]$_.ProcessId) }
        } catch {
        }
    }

    $result = @()
    foreach ($id in $ids) {
        $result += [int]$id
    }
    return $result
}

function Stop-InstalledApp {
    $deadline = (Get-Date).AddSeconds(10)

    do {
        $processIds = @(Get-InstalledAppProcessIds | Where-Object { $_ -and $_ -ne $PID } | Select-Object -Unique)
        if ($processIds.Count -eq 0) {
            return
        }

        Show-InstallStatus "Laufende App-Prozesse werden beendet..."
        foreach ($processId in $processIds) {
            Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
        }
        Start-Sleep -Milliseconds 500
    } while ((Get-Date) -lt $deadline)

    $remaining = @(Get-InstalledAppProcessIds | Where-Object { $_ -and $_ -ne $PID } | Select-Object -Unique)
    if ($remaining.Count -gt 0) {
        throw "Die laufende App konnte nicht vollstaendig beendet werden. Bitte App-Fenster schliessen und install.bat erneut starten."
    }
}

function Remove-InstallDirContentsWithRetry {
    param([string]$TargetDir)

    for ($attempt = 1; $attempt -le 12; $attempt++) {
        try {
            Get-ChildItem -LiteralPath $TargetDir -Force -ErrorAction Stop |
                Where-Object { $_.Name -ne "data" } |
                ForEach-Object {
                    Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction Stop
                }
            return
        } catch {
            if ($attempt -eq 12) {
                throw (
                    "Alte App-Dateien konnten nicht ersetzt werden. " +
                    "Bitte die App, Excel und geoeffnete Einsatzbericht-Dateien schliessen und install.bat erneut starten. " +
                    "Details: " + $_.Exception.Message
                )
            }
            Show-InstallStatus "Alte Dateien sind noch gesperrt. Neuer Versuch $attempt/12..."
            Stop-InstalledApp
            Start-Sleep -Seconds 1
        }
    }
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
    Show-InstallStatus "Python wird gesucht..."
    if (Resolve-PythonCommand) {
        Show-InstallStatus "Python wurde gefunden."
        return
    }

    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Show-InstallStatus "Python wurde nicht gefunden. Installation ueber winget wird versucht..."
        winget install --id Python.Python.3.12 -e --scope user --silent --accept-package-agreements --accept-source-agreements
        Refresh-ProcessPath
        Start-Sleep -Seconds 2
        if (Resolve-PythonCommand) {
            Show-InstallStatus "Python wurde installiert."
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
    Show-InstallStatus "Virtuelle Python-Umgebung wird erstellt..."
    Invoke-BasePython -ArgsList @("-m", "venv", $venvDir)

    if (-not (Test-Path -LiteralPath $venvPython)) {
        throw "Virtuelle Python-Umgebung konnte nicht erstellt werden: $venvDir"
    }

    Show-InstallStatus "pip wird aktualisiert..."
    & $venvPython -m pip install --upgrade pip
    if ($LASTEXITCODE -ne 0) {
        throw "pip konnte nicht aktualisiert werden."
    }

    Show-InstallStatus "Python-Abhaengigkeiten werden installiert. Das kann einige Minuten dauern..."
    & $venvPython -m pip install -r $requirementsPath
    if ($LASTEXITCODE -ne 0) {
        throw "Python-Abhaengigkeiten konnten nicht installiert werden."
    }
}

function Copy-SourcePayload {
    param([string]$SourceDir)

    Show-InstallStatus "Quelldateien werden kopiert..."
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
    if ((Test-Path -LiteralPath $dataSource) -and (-not $script:preserveExistingAppData)) {
        Show-InstallStatus "Leere Datenvorlage wird vorbereitet..."
        Copy-Item -LiteralPath $dataSource -Destination (Join-Path $installDir "data") -Recurse -Force
    } elseif ($script:preserveExistingAppData) {
        Show-InstallStatus "Vorhandener Datenordner bleibt erhalten."
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
$ErrorActionPreference = "Stop"

$installDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$launcherPath = Join-Path $installDir "run_app.exe"
$venvPython = Join-Path $installDir ".venv\Scripts\python.exe"
$sourceLauncher = Join-Path $installDir "run_app.py"

$env:STREAMLIT_BROWSER_GATHER_USAGE_STATS = "false"
$env:STREAMLIT_SERVER_HEADLESS = "true"
$env:STREAMLIT_SERVER_SHOW_EMAIL_PROMPT = "false"
$env:STREAMLIT_GLOBAL_DEVELOPMENT_MODE = "false"

function Show-LaunchError {
    param([string]$Message)
    try {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show($Message, "Startfehler") | Out-Null
    } catch {
        Write-Host $Message
    }
}

try {
    if (Test-Path -LiteralPath $launcherPath) {
        Start-Process -FilePath $launcherPath -WorkingDirectory $installDir | Out-Null
        exit 0
    }
    if ((Test-Path -LiteralPath $venvPython) -and (Test-Path -LiteralPath $sourceLauncher)) {
        Start-Process -FilePath $venvPython -ArgumentList "`"$sourceLauncher`"" -WorkingDirectory $installDir -WindowStyle Hidden | Out-Null
        exit 0
    }
    Show-LaunchError "Die App wurde nicht vollstaendig installiert."
    exit 1
} catch {
    Show-LaunchError ("Die App konnte nicht gestartet werden.`n`nDetails:`n" + $_.Exception.Message)
    exit 1
}
'@

    Set-Content -LiteralPath $launchScriptPath -Value $launchScript -Encoding UTF8
    Unblock-File -LiteralPath $launchScriptPath -ErrorAction SilentlyContinue
}

function Write-LaunchWrapper {
    $command = "`"$powershellPath`" -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$launchScriptPath`""
    $escapedCommand = $command.Replace('"', '""')
    $wrapperScript = @"
Set shell = CreateObject("WScript.Shell")
shell.Run "$escapedCommand", 0, False
"@

    Set-Content -LiteralPath $launchWrapperPath -Value $wrapperScript -Encoding ASCII
    Unblock-File -LiteralPath $launchWrapperPath -ErrorAction SilentlyContinue
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
Show-InstallStep "Installationspaket wird geprueft"
foreach ($candidate in $sourcePayloadCandidates) {
    if ($candidate -and (Test-Path -LiteralPath (Join-Path $candidate $mainScriptName))) {
        $sourcePayloadDir = $candidate
        break
    }
}
if (-not $sourcePayloadDir) {
    $sourcePayloadDir = $scriptRoot
}

Show-InstallStep "Installationsordner werden vorbereitet"
New-Item -ItemType Directory -Path $installRoot -Force | Out-Null
New-Item -ItemType Directory -Path $startMenuDir -Force | Out-Null

if (Test-Path -LiteralPath $installDir) {
    $script:preserveExistingAppData = Test-Path -LiteralPath (Join-Path $installDir "data")
    Show-InstallStatus "Laufende App wird beendet..."
    Stop-InstalledApp
    Show-InstallStatus "Alte App-Dateien werden ersetzt..."
    Remove-InstallDirContentsWithRetry -TargetDir $installDir
}
New-Item -ItemType Directory -Path $installDir -Force | Out-Null

Show-InstallStep "App-Dateien werden kopiert"
if ($useBundledPayload) {
    Show-InstallStatus "Gebundene App-Dateien werden kopiert..."
    Unblock-Tree -RootPath $bundledPayloadDir
    if ($script:preserveExistingAppData) {
        Get-ChildItem -LiteralPath $bundledPayloadDir -Force |
            Where-Object { $_.Name -ne "data" } |
            ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $installDir -Recurse -Force }
    } else {
        Copy-Item -Path (Join-Path $bundledPayloadDir "*") -Destination $installDir -Recurse -Force
    }
    Unblock-Tree -RootPath $installDir
} else {
    Unblock-Tree -RootPath $sourcePayloadDir
    Copy-SourcePayload -SourceDir $sourcePayloadDir
    Unblock-Tree -RootPath $installDir
    Show-InstallStep "Python-Umgebung wird vorbereitet"
    Install-SourceRuntime
}

if ($useBundledPayload) {
    Show-InstallStep "Mitgelieferte App-Runtime wird vorbereitet"
    Show-InstallStatus "Keine Python-Paketinstallation notwendig."
}

Show-InstallStep "Startskript wird erstellt"
Write-LaunchScript
Write-LaunchWrapper

Show-InstallStep "Verknuepfungen werden erstellt"
$wsh = New-Object -ComObject WScript.Shell
foreach ($shortcutPath in @($desktopShortcut, $startMenuShortcut)) {
    $shortcut = $wsh.CreateShortcut($shortcutPath)
    $shortcut.WorkingDirectory = $installDir
    if (Test-Path -LiteralPath $launcherPath) {
        $shortcut.TargetPath = $launcherPath
        $shortcut.Arguments = ""
        $shortcut.IconLocation = $launcherPath
    } elseif (Test-Path -LiteralPath $wscriptPath) {
        $shortcut.TargetPath = $wscriptPath
        $shortcut.Arguments = "`"$launchWrapperPath`""
    } else {
        $shortcut.TargetPath = $powershellPath
        $shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$launchScriptPath`""
    }
    $shortcut.Save()
}

Show-InstallStep "Installation wird abgeschlossen"
Complete-InstallProgress
Show-InfoMessage -Title "Installation abgeschlossen" -Message (
    "Die App wurde erfolgreich installiert.`n`n" +
    "Du kannst sie jetzt ueber den Desktop oder das Startmenue starten."
)
