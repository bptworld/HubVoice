param(
  [ValidateSet("config", "compile")]
  [string]$Action = "compile"
)

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$venvScripts = Join-Path $repoRoot ".envs\runtime\Scripts"
$esphomeExe = Join-Path $venvScripts "esphome.exe"
$platformioExe = Join-Path $venvScripts "platformio.exe"
$patchScript = Join-Path $repoRoot "patch-generated-web-ui.ps1"
$buildDir = Join-Path $repoRoot ".esphome\build\hubvoice-sat"
$env:PATH = "$venvScripts;$repoRoot;$env:PATH"

if (-not (Test-Path $esphomeExe)) {
  throw "ESPHome was not found at $esphomeExe"
}
if (-not (Test-Path $platformioExe)) {
  throw "PlatformIO was not found at $platformioExe"
}

function Invoke-GeneratedSourcePatch {
  & $esphomeExe compile hubvoice-sat.yaml --only-generate
  if ($LASTEXITCODE -ne 0) {
    throw "ESPHome source generation failed"
  }

  & powershell -NoProfile -ExecutionPolicy Bypass -File $patchScript
  if ($LASTEXITCODE -ne 0) {
    throw "Generated web UI patch failed"
  }
}

Push-Location $repoRoot
try {
  if ($Action -eq "config") {
    Invoke-GeneratedSourcePatch
  } else {
    Invoke-GeneratedSourcePatch
    & $platformioExe run -d $buildDir -e hubvoice-sat
    if ($LASTEXITCODE -ne 0) {
      throw "PlatformIO compile failed"
    }
  }
} finally {
  Pop-Location
}
