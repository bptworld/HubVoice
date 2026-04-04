param()

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

function Assert-FileExists([string]$path, [string]$message) {
  if (-not (Test-Path $path)) {
    throw $message
  }
}

function Assert-Contains([string]$text, [string]$pattern, [string]$message) {
  if ($text -notmatch $pattern) {
    throw $message
  }
}

function Assert-NotContains([string]$text, [string]$pattern, [string]$message) {
  if ($text -match $pattern) {
    throw $message
  }
}

$setupPagePath = Join-Path $repoRoot "_live_setup_page.html"
$controlPagePath = Join-Path $repoRoot "control.html"
$setupWebPath = Join-Path $repoRoot "setup-web.ps1"
$runtimePath = Join-Path $repoRoot "hubvoice-runtime.py"
$setupCsprojPath = Join-Path $repoRoot "setup-launcher\HubVoiceSatSetup.csproj"
$runtimeBuildPath = Join-Path $repoRoot "build-runtime-exe.ps1"

Assert-FileExists $setupPagePath "Missing canonical setup page source: $setupPagePath"
Assert-FileExists $controlPagePath "Missing canonical control page source: $controlPagePath"
Assert-FileExists $setupWebPath "Missing setup server script: $setupWebPath"
Assert-FileExists $runtimePath "Missing runtime script: $runtimePath"
Assert-FileExists $setupCsprojPath "Missing setup launcher project: $setupCsprojPath"
Assert-FileExists $runtimeBuildPath "Missing runtime build script: $runtimeBuildPath"

$setupWeb = Get-Content -Path $setupWebPath -Raw -Encoding UTF8
$runtime = Get-Content -Path $runtimePath -Raw -Encoding UTF8
$setupCsproj = Get-Content -Path $setupCsprojPath -Raw -Encoding UTF8
$runtimeBuild = Get-Content -Path $runtimeBuildPath -Raw -Encoding UTF8

# setup-web.ps1 must load the canonical setup HTML file and not carry an inline full page.
Assert-Contains $setupWeb '_live_setup_page\.html' "setup-web.ps1 must load _live_setup_page.html"
Assert-Contains $setupWeb 'Get-Content\s+-Path\s+\$setupPagePath\s+-Raw' "setup-web.ps1 must read setup page from file"
$setupInlineHeredocMarker = '$html = @' + [char]39
if ($setupWeb.Contains($setupInlineHeredocMarker)) {
  throw "setup-web.ps1 contains inline heredoc HTML; use _live_setup_page.html instead"
}
if (($setupWeb -like "*<!doctype html*") -and ($setupWeb -like "*<html*")) {
  throw "setup-web.ps1 contains inline HTML; use _live_setup_page.html instead"
}

# hubvoice-runtime.py must load the canonical control page file and not carry an inline full page blob.
Assert-Contains $runtime 'CONTROL_PAGE_CANDIDATES' "hubvoice-runtime.py must define CONTROL_PAGE_CANDIDATES"
Assert-Contains $runtime 'candidate\.read_text\(encoding="utf-8"\)' "hubvoice-runtime.py must read control.html from file"
Assert-NotContains $runtime 'return\s+"""<!doctype html' "hubvoice-runtime.py contains inline triple-quoted control page"
Assert-NotContains $runtime '\.shell\s*\{' "hubvoice-runtime.py appears to contain embedded control page CSS"

# Build and packaging must include canonical page files.
Assert-Contains $setupCsproj '_live_setup_page\.html' "HubVoiceSatSetup.csproj must embed _live_setup_page.html"
Assert-Contains $runtimeBuild '--add-data",\s+"\$controlPagePath;\.' "build-runtime-exe.ps1 must bundle control.html"

Write-Host "Frontend source-of-truth check passed."
