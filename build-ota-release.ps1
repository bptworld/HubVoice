param()

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$verifyFrontendScript = Join-Path $repoRoot "verify-single-frontend-source.ps1"
$venvScripts = Join-Path $repoRoot ".envs\runtime\Scripts"
$runtimePython = Join-Path $venvScripts "python.exe"
$buildScript = Join-Path $repoRoot "build-hubvoice-sat.ps1"
$secretsPath = Join-Path $repoRoot "secrets.yaml"
$verifyFirmwareBinsScript = Join-Path $repoRoot "verify-firmware-bins.ps1"
$setupProject = Join-Path $repoRoot "setup-launcher\HubVoiceSatSetup.csproj"
$setupBuildDir = Join-Path $repoRoot "build\HubVoiceSatSetup"
$setupExeSource = Join-Path $setupBuildDir "HubVoiceSatSetup.exe"
$runtimeBuildScript = Join-Path $repoRoot "build-runtime-exe.ps1"
$runtimeBuildDir = Join-Path $repoRoot "build\HubVoiceRuntime"
$runtimeExeSource = Join-Path $runtimeBuildDir "HubVoiceRuntime.exe"
$env:PATH = "$venvScripts;$repoRoot;$env:PATH"

if (-not (Test-Path $verifyFrontendScript)) {
  throw "Frontend verification script not found at $verifyFrontendScript"
}
if (-not (Test-Path $verifyFirmwareBinsScript)) {
  throw "Firmware verification script not found at $verifyFirmwareBinsScript"
}

& powershell -NoProfile -ExecutionPolicy Bypass -File $verifyFrontendScript
if ($LASTEXITCODE -ne 0) {
  throw "Frontend source-of-truth verification failed"
}

function Get-YamlValue {
  param(
    [string]$Path,
    [string]$Key
  )

  $pattern = "^\s*${Key}:\s*""?(.*?)""?\s*$"
  $match = Select-String -Path $Path -Pattern $pattern | Select-Object -First 1
  if (-not $match) {
    return $null
  }
  return $match.Matches[0].Groups[1].Value.Trim()
}

function Get-SecretValue {
  param(
    [string]$Path,
    [string]$Key
  )

  if (-not (Test-Path $Path)) {
    return $null
  }

  $pattern = "^\s*${Key}:\s*""?(.*?)""?\s*$"
  $match = Select-String -Path $Path -Pattern $pattern | Select-Object -First 1
  if (-not $match) {
    return $null
  }

  return $match.Matches[0].Groups[1].Value.Trim()
}

function Prepare-ReleaseSecrets {
  param(
    [string]$Path
  )

  $state = @{
    path = $Path
    hadFile = (Test-Path $Path)
    original = $null
    sanitized = $false
  }

  if ($state.hadFile) {
    $state.original = Get-Content -Path $Path -Raw -Encoding UTF8
  }

  $ssid = Get-SecretValue -Path $Path -Key "wifi_ssid"
  $password = Get-SecretValue -Path $Path -Key "wifi_password"

  if ([string]::IsNullOrWhiteSpace($ssid) -and [string]::IsNullOrWhiteSpace($password)) {
    return $state
  }

  $sanitized = @(
    'wifi_ssid: ""'
    'wifi_password: ""'
  )
  Set-Content -Path $Path -Value $sanitized -Encoding UTF8
  $state.sanitized = $true
  Write-Host "Temporarily sanitized secrets.yaml for release build safety."
  return $state
}

function Restore-ReleaseSecrets {
  param(
    [hashtable]$State
  )

  if (-not $State -or -not $State.sanitized) {
    return
  }

  if ($State.hadFile) {
    Set-Content -Path $State.path -Value ([string]$State.original) -Encoding UTF8
  } else {
    Remove-Item -Path $State.path -ErrorAction SilentlyContinue
  }
  Write-Host "Restored local secrets.yaml after release build."
}

function Invoke-ESPHome {
  param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Arguments
  )

  if (-not (Test-Path $runtimePython)) {
    throw "Runtime Python was not found at $runtimePython"
  }

  & $runtimePython -m esphome @Arguments
}

$releaseSecretsState = Prepare-ReleaseSecrets -Path $secretsPath

try {

$yamlPath = Join-Path $repoRoot "hubvoice-sat.yaml"
$version = Get-YamlValue -Path $yamlPath -Key "firmware_version"
$deviceName = Get-YamlValue -Path $yamlPath -Key "device_name"
$fphYamlPath = Join-Path $repoRoot "hubvoice-sat-fph.yaml"
$fphVersion = Get-YamlValue -Path $fphYamlPath -Key "firmware_version"
$fphDeviceName = Get-YamlValue -Path $fphYamlPath -Key "device_name"
$ld2410YamlPath = Join-Path $repoRoot "hubvoice-sat-fph-ld2410.yaml"
$ld2410Version = Get-YamlValue -Path $ld2410YamlPath -Key "firmware_version"
$ld2410DeviceName = Get-YamlValue -Path $ld2410YamlPath -Key "device_name"
if (-not $version) {
  $version = (Get-Date -Format "yyyy.MM.dd.HHmm")
}
if (-not $deviceName) {
  $deviceName = "hubvoice-sat"
}
if (-not $fphVersion) {
  $fphVersion = $version
}
if (-not $fphDeviceName) {
  $fphDeviceName = "hubvoice-sat-fph"
}
if (-not $ld2410Version) {
  $ld2410Version = $fphVersion
}
if (-not $ld2410DeviceName) {
  $ld2410DeviceName = "hubvoice-sat-fph-ld2410"
}

$buildDirs = @(
  (Join-Path $repoRoot ".esphome\build\$deviceName\.pioenvs\$deviceName"),
  (Join-Path $repoRoot ".esphome\build\$deviceName\.pio\build\$deviceName")
)

$fphBuildDirs = @(
  (Join-Path $repoRoot ".esphome\build\$fphDeviceName\.pioenvs\$fphDeviceName"),
  (Join-Path $repoRoot ".esphome\build\$fphDeviceName\.pio\build\$fphDeviceName"),
  (Join-Path $repoRoot ".esphome\build\sat-kitchen\.pioenvs\sat-kitchen")
)

$ld2410BuildDirs = @(
  (Join-Path $repoRoot ".esphome\build\$ld2410DeviceName\.pioenvs\$ld2410DeviceName"),
  (Join-Path $repoRoot ".esphome\build\$ld2410DeviceName\.pio\build\$ld2410DeviceName")
)

$sourceOtaBin = $null
$sourceFactoryBin = $null
$fphSourceOtaBin = $null
$fphSourceFactoryBin = $null
$ld2410SourceOtaBin = $null
$ld2410SourceFactoryBin = $null

foreach ($candidateDir in $buildDirs) {
  $candidateOtaBin = Join-Path $candidateDir "firmware.ota.bin"
  $candidateFactoryBin = Join-Path $candidateDir "firmware.factory.bin"
  if ((Test-Path $candidateOtaBin) -and (Test-Path $candidateFactoryBin)) {
    $sourceOtaBin = $candidateOtaBin
    $sourceFactoryBin = $candidateFactoryBin
    break
  }
}

foreach ($candidateDir in $fphBuildDirs) {
  $candidateOtaBin = Join-Path $candidateDir "firmware.ota.bin"
  $candidateFactoryBin = Join-Path $candidateDir "firmware.factory.bin"
  if ((Test-Path $candidateOtaBin) -and (Test-Path $candidateFactoryBin)) {
    $fphSourceOtaBin = $candidateOtaBin
    $fphSourceFactoryBin = $candidateFactoryBin
    break
  }
}

foreach ($candidateDir in $ld2410BuildDirs) {
  $candidateOtaBin = Join-Path $candidateDir "firmware.ota.bin"
  $candidateFactoryBin = Join-Path $candidateDir "firmware.factory.bin"
  if ((Test-Path $candidateOtaBin) -and (Test-Path $candidateFactoryBin)) {
    $ld2410SourceOtaBin = $candidateOtaBin
    $ld2410SourceFactoryBin = $candidateFactoryBin
    break
  }
}

if (-not $sourceOtaBin -or -not $sourceFactoryBin -or -not (Test-Path $sourceOtaBin) -or -not (Test-Path $sourceFactoryBin)) {
  if (-not (Test-Path $buildScript)) {
    throw "Build script was not found at $buildScript"
  }
  Push-Location $repoRoot
  try {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $buildScript -Action compile
    if ($LASTEXITCODE -ne 0) {
      throw "Firmware compile failed"
    }
  } finally {
    Pop-Location
  }

  foreach ($candidateDir in $buildDirs) {
    $candidateOtaBin = Join-Path $candidateDir "firmware.ota.bin"
    $candidateFactoryBin = Join-Path $candidateDir "firmware.factory.bin"
    if ((Test-Path $candidateOtaBin) -and (Test-Path $candidateFactoryBin)) {
      $sourceOtaBin = $candidateOtaBin
      $sourceFactoryBin = $candidateFactoryBin
      break
    }
  }
}

if (-not $fphSourceOtaBin -or -not $fphSourceFactoryBin -or -not (Test-Path $fphSourceOtaBin) -or -not (Test-Path $fphSourceFactoryBin)) {
  if (-not (Test-Path $fphYamlPath)) {
    throw "FPH YAML config was not found at $fphYamlPath"
  }
  Push-Location $repoRoot
  try {
    Invoke-ESPHome compile $fphYamlPath
    if ($LASTEXITCODE -ne 0) {
      throw "FPH firmware compile failed"
    }
  } finally {
    Pop-Location
  }

  foreach ($candidateDir in $fphBuildDirs) {
    $candidateOtaBin = Join-Path $candidateDir "firmware.ota.bin"
    $candidateFactoryBin = Join-Path $candidateDir "firmware.factory.bin"
    if ((Test-Path $candidateOtaBin) -and (Test-Path $candidateFactoryBin)) {
      $fphSourceOtaBin = $candidateOtaBin
      $fphSourceFactoryBin = $candidateFactoryBin
      break
    }
  }
}

if (-not $ld2410SourceOtaBin -or -not $ld2410SourceFactoryBin -or -not (Test-Path $ld2410SourceOtaBin) -or -not (Test-Path $ld2410SourceFactoryBin)) {
  if (-not (Test-Path $ld2410YamlPath)) {
    throw "LD2410 YAML config was not found at $ld2410YamlPath"
  }
  Push-Location $repoRoot
  try {
    Invoke-ESPHome compile $ld2410YamlPath
    if ($LASTEXITCODE -ne 0) {
      throw "LD2410 firmware compile failed"
    }
  } finally {
    Pop-Location
  }

  foreach ($candidateDir in $ld2410BuildDirs) {
    $candidateOtaBin = Join-Path $candidateDir "firmware.ota.bin"
    $candidateFactoryBin = Join-Path $candidateDir "firmware.factory.bin"
    if ((Test-Path $candidateOtaBin) -and (Test-Path $candidateFactoryBin)) {
      $ld2410SourceOtaBin = $candidateOtaBin
      $ld2410SourceFactoryBin = $candidateFactoryBin
      break
    }
  }
}

if (-not $sourceOtaBin -or -not (Test-Path $sourceOtaBin)) {
  throw "OTA firmware not found after compile"
}
if (-not $sourceFactoryBin -or -not (Test-Path $sourceFactoryBin)) {
  throw "Factory firmware not found after compile"
}
if (-not $fphSourceOtaBin -or -not (Test-Path $fphSourceOtaBin)) {
  throw "FPH OTA firmware not found after compile"
}
if (-not $fphSourceFactoryBin -or -not (Test-Path $fphSourceFactoryBin)) {
  throw "FPH factory firmware not found after compile"
}
if (-not $ld2410SourceOtaBin -or -not (Test-Path $ld2410SourceOtaBin)) {
  throw "LD2410 OTA firmware not found after compile"
}
if (-not $ld2410SourceFactoryBin -or -not (Test-Path $ld2410SourceFactoryBin)) {
  throw "LD2410 factory firmware not found after compile"
}

& powershell -NoProfile -ExecutionPolicy Bypass -File $verifyFirmwareBinsScript -BinPaths @(
  $sourceOtaBin,
  $sourceFactoryBin,
  $fphSourceOtaBin,
  $fphSourceFactoryBin,
  $ld2410SourceOtaBin,
  $ld2410SourceFactoryBin
)
if ($LASTEXITCODE -ne 0) {
  throw "Firmware binary verification failed"
}

$releaseRoot = Join-Path $repoRoot "releases"
$releaseDir = Join-Path $releaseRoot ("hubvoice-sat-" + $version + "-release")
$releaseZip = $releaseDir + ".zip"
$releaseOtaBin = Join-Path $releaseDir ("hubvoice-sat-" + $version + "-ota.bin")
$releaseFactoryBin = Join-Path $releaseDir ("hubvoice-sat-" + $version + "-factory.bin")
$releaseFphOtaBin = Join-Path $releaseDir ("hubvoice-sat-fph-" + $fphVersion + "-ota.bin")
$releaseFphFactoryBin = Join-Path $releaseDir ("hubvoice-sat-fph-" + $fphVersion + "-factory.bin")
$releaseLd2410OtaBin = Join-Path $releaseDir ("hubvoice-sat-fph-ld2410-" + $ld2410Version + "-ota.bin")
$releaseLd2410FactoryBin = Join-Path $releaseDir ("hubvoice-sat-fph-ld2410-" + $ld2410Version + "-factory.bin")
$releaseSetupExe = Join-Path $releaseDir "HubVoiceSatSetup.exe"
$releaseMainExe = Join-Path $releaseDir "HubVoiceSat.exe"
$releaseRuntimeExe = Join-Path $releaseDir "HubVoiceRuntime.exe"
$releaseSatellitesCsv = Join-Path $releaseDir "satellites.csv"
$releasePiperVoicesDir = Join-Path $releaseDir "piper_voices"
$piperVoicesSourceDir = Join-Path $repoRoot "piper_voices"
$instructionsPath = Join-Path $releaseDir "INSTALL.txt"
$checksumsPath = Join-Path $releaseDir "SHA256SUMS.txt"
$openPagePath = Join-Path $releaseDir "open-upload-page.bat"
$openUsbFlashPath = Join-Path $releaseDir "open-usb-flash-page.bat"
$openUsbWifiSetupPath = Join-Path $releaseDir "open-usb-wifi-setup-page.bat"

if (Test-Path $releaseDir) {
  Remove-Item $releaseDir -Recurse -Force
}
if (Test-Path $releaseZip) {
  Remove-Item $releaseZip -Force
}

New-Item -ItemType Directory -Path $releaseDir -Force | Out-Null
Copy-Item $sourceOtaBin $releaseOtaBin -Force
Copy-Item $sourceFactoryBin $releaseFactoryBin -Force
Copy-Item $fphSourceOtaBin $releaseFphOtaBin -Force
Copy-Item $fphSourceFactoryBin $releaseFphFactoryBin -Force
Copy-Item $ld2410SourceOtaBin $releaseLd2410OtaBin -Force
Copy-Item $ld2410SourceFactoryBin $releaseLd2410FactoryBin -Force

if (-not (Test-Path $setupProject)) {
  throw "Setup launcher project was not found at $setupProject"
}

Push-Location $repoRoot
try {
  & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repoRoot "build-setup-launcher.ps1")
  if ($LASTEXITCODE -ne 0) {
    throw "Setup launcher publish failed"
  }
} finally {
  Pop-Location
}

Copy-Item $setupExeSource $releaseSetupExe -Force
Copy-Item $setupExeSource $releaseMainExe -Force

# Build and package standalone runtime executable (no system Python required).
if (-not (Test-Path $runtimeBuildScript)) {
  throw "Runtime build script was not found at $runtimeBuildScript"
}

Push-Location $repoRoot
try {
  & powershell -NoProfile -ExecutionPolicy Bypass -File $runtimeBuildScript
  if ($LASTEXITCODE -ne 0) {
    throw "Runtime executable build failed"
  }
} finally {
  Pop-Location
}

if (-not (Test-Path $runtimeExeSource)) {
  throw "Standalone runtime executable not found at $runtimeExeSource"
}

Copy-Item $runtimeExeSource $releaseRuntimeExe -Force

if (-not (Test-Path $piperVoicesSourceDir)) {
  throw "Required Piper voices directory not found at $piperVoicesSourceDir"
}
Copy-Item $piperVoicesSourceDir $releasePiperVoicesDir -Recurse -Force

# First-run friendly starter file; user fills this in Setup UI.
Set-Content -Path $releaseSatellitesCsv -Value "" -Encoding UTF8

$now = Get-Date
(Get-Item $releaseOtaBin).LastWriteTime = $now
(Get-Item $releaseFactoryBin).LastWriteTime = $now
(Get-Item $releaseFphOtaBin).LastWriteTime = $now
(Get-Item $releaseFphFactoryBin).LastWriteTime = $now
(Get-Item $releaseLd2410OtaBin).LastWriteTime = $now
(Get-Item $releaseLd2410FactoryBin).LastWriteTime = $now
(Get-Item $releaseSetupExe).LastWriteTime = $now
(Get-Item $releaseMainExe).LastWriteTime = $now
(Get-Item $releaseRuntimeExe).LastWriteTime = $now
(Get-Item $releaseSatellitesCsv).LastWriteTime = $now

$instructions = @"
HubVoiceSat Release Package
===========================

Files in this folder:
- HubVoiceSatSetup.exe         Windows setup app â€” launches setup page, starts runtime, lives in system tray
- HubVoiceSat.exe              Primary launcher name (same app as HubVoiceSatSetup.exe)
- HubVoiceRuntime.exe          Standalone runtime (Python bundled, no separate install required)
- satellites.csv               Satellite list (starts empty; fill in Setup)
- piper_voices\               Bundled TTS voice models required by runtime
- $(Split-Path $releaseFactoryBin -Leaf)    First USB install for HA Voice PE (default model)
- $(Split-Path $releaseOtaBin -Leaf)        OTA update for HA Voice PE (default model)
- $(Split-Path $releaseFphFactoryBin -Leaf) First USB install for FPH Satellite-1
- $(Split-Path $releaseFphOtaBin -Leaf)     OTA update for FPH Satellite-1
- $(Split-Path $releaseLd2410FactoryBin -Leaf) First USB install for FPH Satellite-1 LD2410
- $(Split-Path $releaseLd2410OtaBin -Leaf)     OTA update for FPH Satellite-1 LD2410
- open-usb-flash-page.bat      Opens ESP Web Tools for USB flashing
- open-usb-wifi-setup-page.bat Opens serial Wi-Fi provisioning page (USB)
- open-upload-page.bat         Opens satellite web page for OTA updates

FIRST-TIME SETUP
================
Step 1 â€” Flash the satellite (one time per device):
  a. Connect the satellite by USB.
  b. Run open-usb-flash-page.bat.
  c. In ESP Web Tools, connect to the satellite and choose:
      $(Split-Path $releaseFactoryBin -Leaf) (HA Voice PE)
      OR
      $(Split-Path $releaseFphFactoryBin -Leaf) (FPH Satellite-1)
      OR
      $(Split-Path $releaseLd2410FactoryBin -Leaf) (FPH Satellite-1 LD2410)
  d. For HA Voice PE USB onboarding (recommended):
      - After flashing, click the menu (three dots) on the device card.
      - Choose Configure Wi-Fi and enter your home Wi-Fi credentials.
  e. For FPH Satellite-1 fallback onboarding:
      - Join setup Wi-Fi: HV FPH 192.168.4.1
      - Open http://192.168.4.1/ and enter home Wi-Fi credentials.
  f. After the satellite joins home Wi-Fi, open http://<satellite-ip>:8080/
     and set Satellite Name
     to something like Living Room, Kitchen, or Office.

Step 2 â€” Configure and launch:
  Run HubVoiceSat.exe (or HubVoiceSatSetup.exe). It will:
  - Start HubVoiceRuntime.exe automatically
  - Open the setup page in your browser
  - Sit in the system tray while the runtime is active

No separate Python installation is required on the user PC.

Naming multiple satellites:
  Each satellite gets a unique hostname from its MAC address automatically.
  After flashing, set a room label per device from its web page.

OTA UPDATE (later firmware updates)
====================================
1. Make sure the satellite is powered on and on Wi-Fi.
2. Run open-upload-page.bat and enter the satellite IP.
3. In the satellite web page, choose the OTA update option.
4. Select the OTA file matching your hardware:
     - $(Split-Path $releaseOtaBin -Leaf) (HA Voice PE)
     - $(Split-Path $releaseFphOtaBin -Leaf) (FPH Satellite-1)
     - $(Split-Path $releaseLd2410OtaBin -Leaf) (FPH Satellite-1 LD2410)
   Then upload it.
5. Wait for the satellite to reboot.

If you do not know the satellite IP, check your router's client list.
"@
Set-Content -Path $instructionsPath -Value $instructions -Encoding UTF8

$checksums = @(
  (Get-FileHash $releaseSetupExe -Algorithm SHA256),
  (Get-FileHash $releaseRuntimeExe -Algorithm SHA256),
  (Get-FileHash $releaseFactoryBin -Algorithm SHA256),
  (Get-FileHash $releaseOtaBin -Algorithm SHA256),
  (Get-FileHash $releaseFphFactoryBin -Algorithm SHA256),
  (Get-FileHash $releaseFphOtaBin -Algorithm SHA256),
  (Get-FileHash $releaseLd2410FactoryBin -Algorithm SHA256),
  (Get-FileHash $releaseLd2410OtaBin -Algorithm SHA256)
) | ForEach-Object { "{0} *{1}" -f $_.Hash, (Split-Path $_.Path -Leaf) }
Set-Content -Path $checksumsPath -Value $checksums -Encoding ASCII

$openUsbFlashBat = @"
@echo off
start "" "https://web.esphome.io/"
"@
Set-Content -Path $openUsbFlashPath -Value $openUsbFlashBat -Encoding ASCII

$openUsbWifiSetupBat = @"
@echo off
start "" "https://web.esphome.io/"
"@
Set-Content -Path $openUsbWifiSetupPath -Value $openUsbWifiSetupBat -Encoding ASCII

$openPageBat = @"
@echo off
setlocal
set /p DEVICE_IP=Enter satellite IP address: 
if not defined DEVICE_IP exit /b 1
start "" "http://%DEVICE_IP%:8080/"
"@
Set-Content -Path $openPagePath -Value $openPageBat -Encoding ASCII

Compress-Archive -Path (Join-Path $releaseDir "*") -DestinationPath $releaseZip -Force

Write-Host ""
Write-Host "Created combined release:"
Write-Host "  $releaseDir"
Write-Host "Zip package:"
Write-Host "  $releaseZip"
} finally {
  Restore-ReleaseSecrets -State $releaseSecretsState
}
