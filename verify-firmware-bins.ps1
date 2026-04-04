param(
  [string[]]$BinPaths
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$secretsPath = Join-Path $repoRoot "secrets.yaml"
$satellitesPath = Join-Path $repoRoot "satellites.csv"
$setupJsonPath = Join-Path $repoRoot "hubvoice-sat-setup.json"
$rgCommand = Get-Command rg -ErrorAction SilentlyContinue

if (-not $rgCommand) {
  throw "ripgrep (rg) is required for verify-firmware-bins.ps1"
}

if (-not $BinPaths -or $BinPaths.Count -eq 0) {
  throw "Provide at least one firmware .bin path to verify"
}

$resolvedBinPaths = @()
foreach ($binPath in $BinPaths) {
  foreach ($candidatePath in ([string]$binPath -split ',')) {
    $trimmedPath = $candidatePath.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmedPath)) {
      continue
    }
    if (-not (Test-Path $trimmedPath)) {
      throw "Firmware binary not found at $trimmedPath"
    }
    $resolvedBinPaths += (Resolve-Path $trimmedPath).Path
  }
}

function Add-SensitivePattern {
  param(
    [System.Collections.Generic.HashSet[string]]$PatternSet,
    [string]$Value
  )

  $candidate = [string]$Value
  if ([string]::IsNullOrWhiteSpace($candidate)) {
    return
  }

  $trimmed = $candidate.Trim()
  if ($trimmed.Length -lt 4) {
    return
  }

  [void]$PatternSet.Add($trimmed)

  if ($trimmed -match '^https?://') {
    try {
      $uri = [Uri]$trimmed
      $host = [string]$uri.Host
      $isLoopbackHost = $host -in @('127.0.0.1', 'localhost', '::1')
      if (-not $isLoopbackHost -and -not [string]::IsNullOrWhiteSpace($host)) {
        [void]$PatternSet.Add($host)
      }
      if (-not $isLoopbackHost -and -not $uri.IsDefaultPort -and -not [string]::IsNullOrWhiteSpace($uri.Authority)) {
        [void]$PatternSet.Add($uri.Authority)
      }
    } catch {
    }
  }
}

function Get-SecretValue {
  param(
    [string]$Path,
    [string]$Key
  )

  if (-not (Test-Path $Path)) {
    return $null
  }

  $pattern = '^\s*{0}:\s*"?(.*?)"?\s*$' -f [regex]::Escape($Key)
  $match = Select-String -Path $Path -Pattern $pattern | Select-Object -First 1
  if (-not $match) {
    return $null
  }

  return $match.Matches[0].Groups[1].Value.Trim()
}

$patternSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)

Add-SensitivePattern -PatternSet $patternSet -Value (Get-SecretValue -Path $secretsPath -Key "wifi_ssid")
Add-SensitivePattern -PatternSet $patternSet -Value (Get-SecretValue -Path $secretsPath -Key "wifi_password")

if (Test-Path $satellitesPath) {
  $satelliteLines = Get-Content -Path $satellitesPath | Where-Object { $_ -and -not $_.Trim().StartsWith("#") }
  foreach ($line in $satelliteLines) {
    foreach ($part in ($line -split ',')) {
      Add-SensitivePattern -PatternSet $patternSet -Value $part
    }
  }
}

if (Test-Path $setupJsonPath) {
  $setupObject = Get-Content -Path $setupJsonPath -Raw | ConvertFrom-Json
  foreach ($property in $setupObject.PSObject.Properties) {
    Add-SensitivePattern -PatternSet $patternSet -Value ([string]$property.Value)
  }
}

$findings = New-Object System.Collections.Generic.List[string]
foreach ($pattern in ($patternSet | Sort-Object)) {
  foreach ($binPath in $resolvedBinPaths) {
    & $rgCommand.Source -a -F -l -- $pattern $binPath *> $null
    if ($LASTEXITCODE -eq 0) {
      $findings.Add("$pattern => $binPath")
    }
  }
}

if ($findings.Count -gt 0) {
  $details = $findings | ForEach-Object { "- $_" }
  throw (@(
    "Refusing to use firmware binaries because local sensitive values were found inside them.",
    "Checked binaries:",
    ($resolvedBinPaths | ForEach-Object { "- $_" }),
    "Matches:",
    $details
  ) -join [Environment]::NewLine)
}

Write-Host "Firmware binary scan passed. No local sensitive values were found."
