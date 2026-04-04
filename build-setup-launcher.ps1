param()

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$verifyFrontendScript = Join-Path $repoRoot "verify-single-frontend-source.ps1"
$projectPath = Join-Path $repoRoot "setup-launcher\HubVoiceSatSetup.csproj"
$outputPath = Join-Path $repoRoot "build\HubVoiceSatSetup"
$rootExePath = Join-Path $repoRoot "HubVoiceSatSetup.exe"
$rootMainExePath = Join-Path $repoRoot "HubVoiceSat.exe"
$publishedExePath = Join-Path $outputPath "HubVoiceSatSetup.exe"
$yamlPath = Join-Path $repoRoot "hubvoice-sat.yaml"

function Get-YamlValue {
  param(
    [string]$Path,
    [string]$Key
  )

  if (-not (Test-Path $Path)) {
    return $null
  }

  $escapedKey = [regex]::Escape($Key)
  $pattern = '^\s*' + $escapedKey + ':\s*"?(.*?)"?\s*$'
  $match = Select-String -Path $Path -Pattern $pattern | Select-Object -First 1
  if (-not $match) {
    return $null
  }

  return $match.Matches[0].Groups[1].Value.Trim()
}

function Get-SetupAssemblyVersion([string]$rawVersion) {
  if ($rawVersion -and $rawVersion -match '^(\d+)\.(\d+)\.(\d+)\.(\d+)$') {
    return $rawVersion
  }
  return "1.0.0.0"
}

$firmwareVersion = Get-YamlValue -Path $yamlPath -Key "firmware_version"
$setupVersion = Get-SetupAssemblyVersion $firmwareVersion

if (-not (Test-Path $verifyFrontendScript)) {
  throw "Frontend verification script not found at $verifyFrontendScript"
}

& powershell -NoProfile -ExecutionPolicy Bypass -File $verifyFrontendScript
if ($LASTEXITCODE -ne 0) {
  throw "Frontend source-of-truth verification failed"
}

if (Test-Path $outputPath) {
  Remove-Item $outputPath -Recurse -Force
}

$publishArgs = @(
  "publish"
  $projectPath
  "-c"
  "Release"
  "-o"
  $outputPath
  "-p:Version=$setupVersion"
  "-p:AssemblyVersion=$setupVersion"
  "-p:FileVersion=$setupVersion"
  "-p:InformationalVersion=HubVoiceSatSetup $setupVersion"
)

& dotnet @publishArgs
if ($LASTEXITCODE -ne 0) {
  throw "Setup launcher publish failed"
}

if (-not (Test-Path $publishedExePath)) {
  throw "Published setup launcher was not found at $publishedExePath"
}

Copy-Item $publishedExePath $rootExePath -Force
Copy-Item $publishedExePath $rootMainExePath -Force

Write-Host ""
Write-Host "Created single-file setup launcher:"
Write-Host "  $publishedExePath"
Write-Host "Refreshed repo launcher:"
Write-Host "  $rootExePath"
Write-Host "Refreshed repo primary launcher alias:"
Write-Host "  $rootMainExePath"
Write-Host "Setup EXE version:"
Write-Host "  $setupVersion"
