#!/usr/bin/env powershell
<#
.SYNOPSIS
HubVoiceSat Flash Helper - OTA firmware upload for satellite devices
#>

param(
    [string]$Action = "upload",
    [string]$Device = "",
    [string]$IP = "",
    [string]$Model = ""
)

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$EsphomeVenv = Join-Path $Root ".runtime-venv\Scripts\esphome.exe"
$SatellitesCSV = Join-Path $Root "satellites.csv"

# Select YAML based on hardware model
$YamlMap = @{
    "echos3r" = "hubvoice-sat-echos3r.yaml"
    "fph"     = "hubvoice-sat-fph.yaml"
    "sat1"    = "hubvoice-sat-fph.yaml"
    "sat-1"   = "hubvoice-sat-fph.yaml"
    "default" = "hubvoice-sat.yaml"
}
if (-not $Model) {
    Write-Host "Hardware model (leave blank for default HA Voice PE):"
    Write-Host "  [blank] = HA Voice PE (hubvoice-sat.yaml)"
    Write-Host "  echos3r = M5Stack Atom EchoS3R (hubvoice-sat-echos3r.yaml)"
    Write-Host "  fph     = FPH Satellite-1 (hubvoice-sat-fph.yaml)"
    $Model = Read-Host "Model"
}
$YamlKey = if ($YamlMap.ContainsKey($Model.ToLower())) { $Model.ToLower() } else { "default" }
$YamlFile = Join-Path $Root $YamlMap[$YamlKey]
Write-Host "Using config: $($YamlMap[$YamlKey])"

# Check ESPHome installed
if (-not (Test-Path $EsphomeVenv)) {
    Write-Host "ESPHome is not installed. Run:"
    Write-Host "  .\.runtime-venv\Scripts\python.exe -m pip install esphome==2026.2.4"
    [Environment]::Exit(1)
}

# Load first satellite from CSV
$DefaultDevice = "hubvoice-sat"
$DefaultIP = ""
if (Test-Path $SatellitesCSV) {
    $csv = @(Get-Content $SatellitesCSV | Where-Object {$_ -and -not $_.StartsWith("#")})
    if ($csv.Count -gt 0) {
        $parts = $csv[0].Split(",")
        $first = $parts[0].Trim()
        $firstIP = if ($parts.Count -ge 2) { $parts[1].Trim() } else { "" }
        if ($first) {
            $DefaultDevice = $first
            $DefaultIP = $firstIP
        }
    }
}

# Interactive mode: prompt if no device specified
$TargetDevice = $Device
$TargetIP = $IP

if (-not $TargetDevice) {
    Write-Host "Known satellite names on this PC:"
    if (Test-Path $SatellitesCSV) {
        Get-Content $SatellitesCSV | Where-Object {$_ -and -not $_.StartsWith("#")} | ForEach-Object {
            $parts = $_.Split(",")
            Write-Host "  $($parts[0].Trim())  [$($parts[1].Trim())]"
        }
    } else {
        Write-Host "  (none yet)"
    }
    Write-Host ""
    $TargetDevice = Read-Host "Satellite device name [$DefaultDevice]"
    if (-not $TargetDevice) {
        $TargetDevice = $DefaultDevice
    }
}

# Look up IP from CSV
if (-not $TargetIP -and (Test-Path $SatellitesCSV)) {
    $csv = Get-Content $SatellitesCSV | Where-Object {$_ -and -not $_.StartsWith("#")} | Select-String "^$TargetDevice,"
    if ($csv) {
        $parts = $csv.ToString().Split(",")
        $TargetIP = $parts[1].Trim()
    }
}

if ($TargetIP) {
    Write-Host "Saved OTA IP for $TargetDevice`: $TargetIP"
} else {
    $TargetIP = Read-Host "Satellite IP address (leave blank to try .local)"
    if (-not $TargetIP) {
        $TargetIP = "$TargetDevice.local"
    }
}

Write-Host ""
Write-Host "========================================"
Write-Host "HubVoiceSat Flash Helper"
Write-Host "Device Name: $TargetDevice"
Write-Host "OTA Target: $TargetIP"
Write-Host "========================================"
Write-Host ""

# Execute flash
$esphomeArgs = @(
    "-s", "device_name", $TargetDevice,
    "-s", "friendly_name", $TargetDevice,
    "run",
    $YamlFile,
    "--device", $TargetIP,
    "--no-logs"  # Skip interactive monitoring after upload
)

Write-Host "Running: esphome $($esphomeArgs -join ' ')"
Write-Host ""

& $EsphomeVenv @esphomeArgs

if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "Flash failed with exit code $LASTEXITCODE"
    Read-Host "Press Enter to exit"
    [Environment]::Exit($LASTEXITCODE)
}

Write-Host ""
Write-Host "Flash completed successfully!"
