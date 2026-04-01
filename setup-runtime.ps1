param(
  [switch]$Force
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$venvDir = Join-Path $root ".runtime-venv"
$venvPython = Join-Path $venvDir "Scripts\python.exe"
$venvPip = Join-Path $venvDir "Scripts\pip.exe"
$requirementsPath = Join-Path $root "requirements-runtime.txt"

Write-Host ""
Write-Host "HubVoice Runtime Setup"
Write-Host "======================"
Write-Host ""

# --- Check Python ---
$pythonCmd = $null
foreach ($candidate in @("python", "python3", "py")) {
  try {
    $ver = & $candidate --version 2>&1
    if ($ver -match "Python 3\.(\d+)") {
      $minor = [int]$Matches[1]
      if ($minor -ge 10) {
        $pythonCmd = $candidate
        Write-Host "Found: $ver ($candidate)"
        break
      } else {
        Write-Host "WARN: $ver is too old (need 3.10+). Trying next..."
      }
    }
  } catch {
    # not found, try next
  }
}

if (-not $pythonCmd) {
  Write-Host ""
  Write-Host "ERROR: Python 3.10 or newer is required but was not found."
  Write-Host "       Download from https://www.python.org/downloads/"
  Write-Host "       Make sure to check 'Add Python to PATH' during install."
  exit 1
}

# --- Create or reuse venv ---
if ($Force -and (Test-Path $venvDir)) {
  Write-Host ""
  Write-Host "Removing existing venv (-Force was specified)..."
  Remove-Item $venvDir -Recurse -Force
}

if (-not (Test-Path $venvPython)) {
  Write-Host ""
  Write-Host "Creating virtual environment at .runtime-venv ..."
  & $pythonCmd -m venv $venvDir
  if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Failed to create virtual environment."
    exit 1
  }
  Write-Host "Done."
} else {
  Write-Host ""
  Write-Host "Virtual environment already exists. Updating packages..."
}

# --- Upgrade pip ---
Write-Host ""
Write-Host "Upgrading pip..."
& $venvPython -m pip install --upgrade pip --quiet
if ($LASTEXITCODE -ne 0) {
  Write-Host "WARN: pip upgrade failed — continuing anyway."
}

# --- Install runtime packages ---
Write-Host ""
Write-Host "Installing runtime packages from requirements-runtime.txt..."
Write-Host "(This may take a few minutes on first run.)"
Write-Host ""
& $venvPip install -r $requirementsPath
if ($LASTEXITCODE -ne 0) {
  Write-Host ""
  Write-Host "ERROR: Package installation failed. See output above for details."
  exit 1
}

# --- Verify key imports ---
Write-Host ""
Write-Host "Verifying installed packages..."

$checks = @(
  @{ import = "aioesphomeapi"; label = "aioesphomeapi (satellite communication)" },
  @{ import = "faster_whisper"; label = "faster-whisper (speech recognition)" },
  @{ import = "piper.voice"; label = "piper-tts (text to speech)" },
  @{ import = "av"; label = "av / PyAV (HubMusic MP3 encoding)" },
  @{ import = "numpy"; label = "numpy (audio processing)" },
  @{ import = "sounddevice"; label = "sounddevice (HubMusic desktop audio relay)" }
)

$allOk = $true
foreach ($check in $checks) {
  $result = & $venvPython -c "import $($check.import); print('ok')" 2>&1
  if ($result -match "ok") {
    Write-Host "  OK  $($check.label)"
  } else {
    Write-Host "  FAIL $($check.label)"
    $allOk = $false
  }
}

Write-Host ""
if ($allOk) {
  Write-Host "All packages verified. Runtime is ready."
  Write-Host ""
  Write-Host "Next steps:"
  Write-Host "  1. Make sure piper_voices\ contains a .onnx voice model"
  Write-Host "  2. Configure hubvoice-sat-setup.json (or run HubVoiceSatSetup.exe)"
  Write-Host "  3. Start the runtime: .runtime-venv\Scripts\python.exe hubvoice-runtime.py"
} else {
  Write-Host "One or more packages failed to import. See errors above."
  exit 1
}
