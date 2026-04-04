param()

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$verifyFrontendScript = Join-Path $repoRoot "verify-single-frontend-source.ps1"
$runtimeScript = Join-Path $repoRoot "hubvoice-runtime.py"
$runtimeVenvPython = Join-Path $repoRoot ".envs\runtime\Scripts\python.exe"
$fallbackPython = Join-Path $repoRoot ".envs\main\Scripts\python.exe"
$outputPath = Join-Path $repoRoot "build\HubVoiceRuntime"
$workPath = Join-Path $repoRoot "build\HubVoiceRuntime-work"
$specPath = Join-Path $repoRoot "build\HubVoiceRuntime-spec"
$publishedExePath = Join-Path $outputPath "HubVoiceRuntime.exe"
$rootExePath = Join-Path $repoRoot "HubVoiceRuntime.exe"
$controlPagePath = Join-Path $repoRoot "control.html"

if (-not (Test-Path $runtimeScript)) {
  throw "Runtime script not found at $runtimeScript"
}

if (-not (Test-Path $controlPagePath)) {
  throw "Control page not found at $controlPagePath"
}

if (-not (Test-Path $verifyFrontendScript)) {
  throw "Frontend verification script not found at $verifyFrontendScript"
}

& powershell -NoProfile -ExecutionPolicy Bypass -File $verifyFrontendScript
if ($LASTEXITCODE -ne 0) {
  throw "Frontend source-of-truth verification failed"
}

$pythonExe = if (Test-Path $runtimeVenvPython) { $runtimeVenvPython } elseif (Test-Path $fallbackPython) { $fallbackPython } else { "python" }

if (Test-Path $outputPath) { Remove-Item $outputPath -Recurse -Force }
if (Test-Path $workPath) { Remove-Item $workPath -Recurse -Force }
if (Test-Path $specPath) { Remove-Item $specPath -Recurse -Force }

New-Item -ItemType Directory -Force -Path $outputPath | Out-Null
New-Item -ItemType Directory -Force -Path $workPath | Out-Null
New-Item -ItemType Directory -Force -Path $specPath | Out-Null

& $pythonExe -m pip install --upgrade pip pyinstaller
if ($LASTEXITCODE -ne 0) {
  throw "Failed to install PyInstaller"
}

$pyInstallerArgs = @(
  "-m", "PyInstaller",
  "--noconfirm",
  "--clean",
  "--onefile",
  "--name", "HubVoiceRuntime",
  "--distpath", $outputPath,
  "--workpath", $workPath,
  "--specpath", $specPath,
  "--paths", $repoRoot,
  "--add-data", "$controlPagePath;.",
  $runtimeScript
)

& $pythonExe @pyInstallerArgs
if ($LASTEXITCODE -ne 0) {
  throw "PyInstaller build failed"
}

if (-not (Test-Path $publishedExePath)) {
  throw "Published runtime executable was not found at $publishedExePath"
}

Copy-Item $publishedExePath $rootExePath -Force

Write-Host ""
Write-Host "Created standalone runtime launcher:"
Write-Host "  $publishedExePath"
Write-Host "Refreshed repo runtime launcher:"
Write-Host "  $rootExePath"
