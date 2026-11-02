param(
    [switch]$Uninstall
)

$ErrorActionPreference = "Stop"
$taskName = "Service GUID Processor Azure Token Helper"

if ($Uninstall) {
    $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if ($task) {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        Write-Host "Removed the Azure token helper startup task."
    } else {
        Write-Host "The Azure token helper startup task is not installed."
    }
    return
}

$projectDirectory = $PSScriptRoot
$python = Join-Path $projectDirectory ".venv\Scripts\python.exe"
if (-not (Test-Path $python)) {
    throw "Project Python environment not found. Run run.bat once, then install this startup task."
}

$logDirectory = Join-Path $projectDirectory "results"
New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
$logPath = Join-Path $logDirectory "azure-token-helper.log"
$command = "`"$python`" -m uvicorn app:app --host 127.0.0.1 --port 8000 >> `"$logPath`" 2>&1"
$userId = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$action = New-ScheduledTaskAction `
    -Execute $env:ComSpec `
    -Argument "/c `"$command`"" `
    -WorkingDirectory $projectDirectory
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $userId
$principal = New-ScheduledTaskPrincipal `
    -UserId $userId `
    -LogonType Interactive `
    -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -RestartInterval (New-TimeSpan -Minutes 1) `
    -RestartCount 3 `
    -ExecutionTimeLimit ([TimeSpan]::Zero)

Register-ScheduledTask `
    -TaskName $taskName `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Settings $settings `
    -Description "Starts the local Azure CLI token helper when this user signs in." `
    -Force | Out-Null

Write-Host "Installed the Azure token helper startup task for $userId."
Write-Host "It will start when you sign in to Windows. Keep this PC awake and signed in while using token generation."
Write-Host "Startup output is logged to $logPath."
