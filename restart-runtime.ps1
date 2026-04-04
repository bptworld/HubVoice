[CmdletBinding()]
param(
    [switch]$ShutdownAll
)

$ErrorActionPreference = 'Stop'

$root = $PSScriptRoot
if (-not $root) {
    $root = (Get-Location).Path
}

function Stop-ProcessesByFilter {
    param(
        [string]$Filter,
        [scriptblock]$Predicate,
        [string]$Label
    )

    try {
        Get-CimInstance Win32_Process -Filter $Filter |
            Where-Object $Predicate |
            ForEach-Object {
                try {
                    Stop-Process -Id $_.ProcessId -Force -ErrorAction Stop
                    Write-Host ("Stopped {0} PID={1}" -f $Label, $_.ProcessId)
                } catch { }
            }
    } catch { }
}

function Stop-SetupPortListeners {
    param(
        [int]$BasePort = 8093,
        [int]$Span = 10
    )

    for ($port = $BasePort; $port -lt ($BasePort + $Span); $port++) {
        try {
            $listeners = Get-NetTCPConnection -State Listen -LocalPort $port -ErrorAction SilentlyContinue
            foreach ($listener in @($listeners)) {
                try {
                    $owner = [int]$listener.OwningProcess
                    if ($owner -and $owner -ne $PID) {
                        Stop-Process -Id $owner -Force -ErrorAction Stop
                        Write-Host ("Stopped setup port listener PID={0} on port {1}" -f $owner, $port)
                    }
                } catch { }
            }
        } catch { }
    }
}

if ($ShutdownAll) {
    # Stop setup web hosts started from this workspace.
    Stop-ProcessesByFilter -Filter "Name = 'powershell.exe' OR Name = 'pwsh.exe'" -Label "setup-web" -Predicate {
        ($_.ProcessId -ne $PID) -and ([string]$_.CommandLine -like '*setup-web.ps1*') -and ([string]$_.CommandLine -like "*$root*")
    }

    # Stop launcher executable from this workspace if running.
    Stop-ProcessesByFilter -Filter "Name = 'HubVoiceSatSetup.exe'" -Label "launcher" -Predicate {
        [string]$_.ExecutablePath -like "*$root*"
    }

    # Stop standalone frozen runtime executable from this workspace if running.
    Stop-ProcessesByFilter -Filter "Name = 'HubVoiceRuntime.exe'" -Label "runtime exe" -Predicate {
        [string]$_.ExecutablePath -like "*$root*"
    }

    $setupPort = 8093
    try {
        if ($env:HUBVOICESAT_SETUP_PORT) {
            $candidate = [int]$env:HUBVOICESAT_SETUP_PORT
            if ($candidate -ge 1024 -and $candidate -le 65535) {
                $setupPort = $candidate
            }
        }
    } catch { }

    # Defensive cleanup: remove any stale setup listeners even if process cmdline does not include setup-web.ps1.
    Stop-SetupPortListeners -BasePort $setupPort -Span 10
}

# Always stop runtime Python processes for a clean restart.
Stop-ProcessesByFilter -Filter "Name = 'python.exe' OR Name = 'pythonw.exe'" -Label "runtime python" -Predicate {
    ([string]$_.CommandLine -like '*hubvoice-runtime.py*') -and ([string]$_.CommandLine -like "*$root*")
}

Start-Sleep -Milliseconds 800

$pythonCandidates = @(
    (Join-Path $root '.envs\runtime\Scripts\python.exe'),
    (Join-Path $root '.envs\main\Scripts\python.exe')
)
$pythonExe = $pythonCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $pythonExe) {
    throw "Unable to find python.exe in .envs/runtime or .envs/main under $root"
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

if ($ShutdownAll) {
    $setupScript = Join-Path $root 'setup-web.ps1'
    if (Test-Path $setupScript) {
        $psExe = if (Test-Path "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe") {
            "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
        } else {
            'powershell.exe'
        }

        $setupProc = Start-Process `
            -FilePath $psExe `
            -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $setupScript) `
            -WorkingDirectory $root `
            -WindowStyle Hidden `
            -PassThru

        $setupPort = 8093
        try {
            if ($env:HUBVOICESAT_SETUP_PORT) {
                $candidate = [int]$env:HUBVOICESAT_SETUP_PORT
                if ($candidate -ge 1024 -and $candidate -le 65535) {
                    $setupPort = $candidate
                }
            }
        } catch { }

        $setupDeadline = (Get-Date).AddSeconds(15)
        $setupReady = $false
        while ((Get-Date) -lt $setupDeadline) {
            try {
                if (Test-NetConnection -ComputerName 127.0.0.1 -Port $setupPort -InformationLevel Quiet) {
                    $setupReady = $true
                    break
                }
            } catch { }
            Start-Sleep -Milliseconds 300
        }

        if ($setupReady) {
            Write-Host "Setup page started (PID=$($setupProc.Id)) and is listening on 127.0.0.1:$setupPort"
        } else {
            Write-Warning "Setup page process started (PID=$($setupProc.Id)) but port $setupPort is not reachable yet."
        }
    } else {
        Write-Warning "setup-web.ps1 was not found under $root; setup page was not restarted."
    }
}
