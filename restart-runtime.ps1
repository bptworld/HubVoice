$ErrorActionPreference = 'Stop'

$root = $PSScriptRoot
if (-not $root) {
    $root = (Get-Location).Path
}

# Stop any running hubvoice-runtime Python processes.
Get-CimInstance Win32_Process -Filter "Name = 'python.exe'" |
    Where-Object { $_.CommandLine -like '*hubvoice-runtime.py*' } |
    ForEach-Object {
        try { Stop-Process -Id $_.ProcessId -Force } catch { }
    }

Start-Sleep -Milliseconds 800

$pythonCandidates = @(
    (Join-Path $root '.runtime-venv\Scripts\python.exe'),
    (Join-Path $root '.venv\Scripts\python.exe')
)
$pythonExe = $pythonCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $pythonExe) {
    throw "Unable to find python.exe in .runtime-venv or .venv under $root"
}

$env:PYTHONUNBUFFERED = '1'
$stdoutLog = Join-Path $root 'logs\hubvoice-runtime.log'
$stderrLog = Join-Path $root 'logs\hubvoice-runtime-err.log'

$proc = Start-Process `
    -FilePath $pythonExe `
    -ArgumentList @((Join-Path $root 'hubvoice-runtime.py')) `
    -WorkingDirectory $root `
    -WindowStyle Hidden `
    -RedirectStandardOutput $stdoutLog `
    -RedirectStandardError $stderrLog `
    -PassThru

# Wait up to 20s for control page to come up.
$deadline = (Get-Date).AddSeconds(20)
$ready = $false
while ((Get-Date) -lt $deadline) {
    try {
        $ok = Test-NetConnection -ComputerName 127.0.0.1 -Port 8080 -InformationLevel Quiet
        if ($ok) {
            $ready = $true
            break
        }
    } catch { }
    Start-Sleep -Milliseconds 300
}

if ($ready) {
    Write-Host "HubVoice runtime started (PID=$($proc.Id)) and is listening on 127.0.0.1:8080"
} else {
    Write-Warning "HubVoice runtime started (PID=$($proc.Id)) but 8080 is not reachable yet. Check logs\\hubvoice-runtime-err.log"
}
