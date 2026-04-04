@echo off
setlocal

set "ROOT=%~dp0"
if "%ROOT:~-1%"=="\" set "ROOT=%ROOT:~0,-1%"

echo [HubVoice] Stopping local launcher, setup, and runtime processes...

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$root = [System.IO.Path]::GetFullPath('%ROOT%');" ^
  "$killed = New-Object System.Collections.Generic.HashSet[int];" ^
  "function Stop-MatchingProcess([int]$targetProcessId) { if (-not $targetProcessId) { return }; if ($targetProcessId -eq $PID) { return }; if ($killed.Add($targetProcessId)) { try { Stop-Process -Id $targetProcessId -Force -ErrorAction Stop; Write-Host ('Stopped PID ' + $targetProcessId) } catch { } } };" ^
  "$py = Get-CimInstance Win32_Process -Filter \"Name = 'python.exe' OR Name = 'pythonw.exe'\";" ^
  "foreach ($proc in $py) { $cmd = [string]$proc.CommandLine; if ($cmd -and (($cmd -match 'hubvoice-runtime\\.py') -or ($cmd -match 'setup-web\\.ps1')) -and ($cmd -like ('*' + $root + '*'))) { Stop-MatchingProcess ([int]$proc.ProcessId) } }" ^
  "$pwsh = Get-CimInstance Win32_Process -Filter \"Name = 'powershell.exe' OR Name = 'pwsh.exe'\";" ^
  "foreach ($proc in $pwsh) { $cmd = [string]$proc.CommandLine; if ($cmd -and ($cmd -match 'setup-web\\.ps1') -and ($cmd -like ('*' + $root + '*'))) { Stop-MatchingProcess ([int]$proc.ProcessId) } }" ^
  "$exes = Get-CimInstance Win32_Process | Where-Object { (($_.Name -eq 'HubVoiceRuntime.exe') -or ($_.Name -eq 'HubVoiceSatSetup.exe') -or ($_.Name -eq 'HubVoiceSat.exe')) -and ([string]$_.ExecutablePath -like ('*' + $root + '*')) };" ^
  "foreach ($proc in $exes) { Stop-MatchingProcess ([int]$proc.ProcessId) }" ^
  "foreach ($port in 8080,8093) { try { Get-NetTCPConnection -State Listen -LocalPort $port -ErrorAction SilentlyContinue | ForEach-Object { Stop-MatchingProcess ([int]$_.OwningProcess) } } catch { } }"

if errorlevel 1 (
  echo [HubVoice] Warning: one or more stop operations returned an error.
)

echo [HubVoice] Starting live setup page from source...
set "SETUP_URL=http://127.0.0.1:8093/"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$env:HUBVOICESAT_SUPPRESS_AUTO_OPEN='1';" ^
  "$scriptPath = [System.IO.Path]::Combine('%ROOT%', 'setup-web.ps1');" ^
  "$psi = New-Object System.Diagnostics.ProcessStartInfo;" ^
  "$psi.FileName = 'powershell.exe';" ^
  "$psi.Arguments = '-NoProfile -ExecutionPolicy Bypass -File \"' + $scriptPath + '\"';" ^
  "$psi.WorkingDirectory = '%ROOT%';" ^
  "$psi.UseShellExecute = $false;" ^
  "$psi.CreateNoWindow = $true;" ^
  "[System.Diagnostics.Process]::Start($psi) | Out-Null"
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
  echo [HubVoice] Live setup page launch failed with exit code %EXIT_CODE%.
  exit /b %EXIT_CODE%
)

echo [HubVoice] Waiting for setup page...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$deadline = (Get-Date).AddSeconds(20);" ^
  "$ready = $false;" ^
  "while ((Get-Date) -lt $deadline) { try { $response = Invoke-WebRequest -UseBasicParsing '%SETUP_URL%'; if ($response.StatusCode -ge 200 -and $response.StatusCode -lt 500) { $ready = $true; break } } catch { } Start-Sleep -Milliseconds 300 }" ^
  "if (-not $ready) { exit 1 }"

if errorlevel 1 (
  echo [HubVoice] Setup page did not come up on %SETUP_URL%.
  exit /b 1
)

echo [HubVoice] Opening %SETUP_URL%
start "" "%SETUP_URL%"

echo [HubVoice] Live setup page is refreshed.
exit /b 0