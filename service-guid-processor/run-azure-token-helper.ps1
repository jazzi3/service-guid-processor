$ErrorActionPreference = "Stop"

$projectDirectory = $PSScriptRoot
$python = Join-Path $projectDirectory ".venv\Scripts\python.exe"
$logPath = Join-Path $projectDirectory "results\azure-token-helper.log"
$stdoutPath = Join-Path $projectDirectory "results\azure-token-helper.stdout.log"
$stderrPath = Join-Path $projectDirectory "results\azure-token-helper.stderr.log"

if (-not (Test-Path $python)) {
    throw "Project Python environment not found. Run run.bat once before starting the Azure token helper."
}

New-Item -ItemType Directory -Path (Split-Path $logPath) -Force | Out-Null

while ($true) {
    try {
        "Starting Azure token helper at $(Get-Date -Format o)" | Out-File -FilePath $logPath -Append -Encoding utf8
        $server = Start-Process `
            -FilePath $python `
            -ArgumentList @("-m", "uvicorn", "app:app", "--host", "127.0.0.1", "--port", "8000") `
            -WorkingDirectory $projectDirectory `
            -RedirectStandardOutput $stdoutPath `
            -RedirectStandardError $stderrPath `
            -WindowStyle Hidden `
            -PassThru
        $server.WaitForExit()

        "Azure token helper exited with code $($server.ExitCode) at $(Get-Date -Format o)" | Out-File -FilePath $logPath -Append -Encoding utf8
        foreach ($outputPath in @($stdoutPath, $stderrPath)) {
            if (Test-Path $outputPath) {
                Get-Content $outputPath | Out-File -FilePath $logPath -Append -Encoding utf8
                Clear-Content $outputPath
            }
        }
    } catch {
        "Azure token helper launch failed at $(Get-Date -Format o): $($_.Exception.Message)" | Out-File -FilePath $logPath -Append -Encoding utf8
    }

    Start-Sleep -Seconds 5
}
