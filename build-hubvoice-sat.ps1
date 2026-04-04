param(
  [ValidateSet("config", "compile")]
  [string]$Action = "compile"
)

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$venvScripts = Join-Path $repoRoot ".envs\runtime\Scripts"
$runtimePython = Join-Path $venvScripts "python.exe"
$patchScript = Join-Path $repoRoot "patch-generated-web-ui.ps1"
$buildDir = Join-Path $repoRoot ".esphome\build\hubvoice-sat"
$env:PATH = "$venvScripts;$repoRoot;$env:PATH"

if (-not (Test-Path $runtimePython)) {
  throw "Runtime Python was not found at $runtimePython"
}

function Invoke-ESPHome {
  param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Arguments
  )

  & $runtimePython -m esphome @Arguments
}

function Invoke-PlatformIO {
  param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Arguments
  )

  & $runtimePython -m platformio @Arguments
}

function Invoke-GeneratedSourcePatch {
  Invoke-ESPHome compile hubvoice-sat.yaml --only-generate
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
    $platformioArgs = @("run", "-d", $buildDir, "-e", "hubvoice-sat")
    Invoke-PlatformIO @platformioArgs
    if ($LASTEXITCODE -ne 0) {
      throw "PlatformIO compile failed"
    }
  }
} finally {
  Pop-Location
}
