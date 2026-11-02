param(
    [switch]$Uninstall
)

$ErrorActionPreference = "Stop"

$taskName = "Service GUID Processor Azure Token Helper"
$shortcutName = "Service GUID Processor Azure Token Helper.lnk"
$startupDirectory = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\Startup"
$shortcutPath = Join-Path $startupDirectory $shortcutName

if ($Uninstall) {
    $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if ($task) {
        Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    }
    if (Test-Path $shortcutPath) {
        Remove-Item -LiteralPath $shortcutPath
    }
    Write-Host "Removed the Azure token helper auto-start configuration."
    return
}

$projectDirectory = $PSScriptRoot
$python = Join-Path $projectDirectory ".venv\Scripts\python.exe"
if (-not (Test-Path $python)) {
    throw "Project Python environment not found. Run run.bat once, then install this startup task."
}

$runner = Join-Path $projectDirectory "run-azure-token-helper.ps1"
if (-not (Test-Path $runner)) {
    throw "Azure token helper supervisor not found: $runner"
}

$oldTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($oldTask) {
    Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

$logDirectory = Join-Path $projectDirectory "results"
New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
$logPath = Join-Path $logDirectory "azure-token-helper.log"
$powershell = Join-Path $PSHOME "powershell.exe"
if (-not (Test-Path $startupDirectory)) {
    New-Item -ItemType Directory -Path $startupDirectory -Force | Out-Null
}

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $powershell
$shortcut.Arguments = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$runner`""
$shortcut.WorkingDirectory = $projectDirectory
$shortcut.WindowStyle = 7
$shortcut.Description = "Starts the local Azure CLI token helper at Windows sign-in."
$shortcut.Save()

Write-Host "Installed the Azure token helper startup shortcut for the current Windows user."
Write-Host "It will start when you sign in to Windows and restart the helper if it exits."
Write-Host "Keep this PC awake and signed in while using token generation."
Write-Host "Startup output is logged to $logPath."

$healthy = $false
try {
    $health = Invoke-RestMethod -Uri "http://127.0.0.1:8000/health" -TimeoutSec 2
    $healthy = $health.status -eq "ok"
} catch {
    $healthy = $false
}
if (-not $healthy) {
    Start-Process `
        -FilePath $powershell `
        -ArgumentList $shortcut.Arguments `
        -WorkingDirectory $projectDirectory `
        -WindowStyle Hidden | Out-Null
    Write-Host "Started the Azure token helper."
} else {
    Write-Host "The Azure token helper is already running."
}
