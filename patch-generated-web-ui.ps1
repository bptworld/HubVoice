param()

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$mainCppPath = Join-Path $repoRoot ".esphome\build\hubvoice-sat\src\main.cpp"
$customJsPath = Join-Path $repoRoot "web-ui-custom.js"

if (-not (Test-Path $mainCppPath)) {
  throw "Generated source not found at $mainCppPath"
}

if (-not (Test-Path $customJsPath)) {
  throw "Custom web UI script not found at $customJsPath"
}

$mainCpp = Get-Content $mainCppPath -Raw
$customJs = (Get-Content $customJsPath -Raw).Replace("</script>", "<\/script>")
$html = @"
<!DOCTYPE html><html><head><meta charset=UTF-8><link rel=icon href=data:></head><body><esp-app></esp-app><script src="https://oi.esphome.io/v2/www.js"></script><script>$customJs</script></body></html>
"@

$bytes = [System.Text.Encoding]::UTF8.GetBytes($html)
$bytesText = ($bytes | ForEach-Object { $_.ToString() }) -join ", "
$arrayLine = "const uint8_t ESPHOME_WEBSERVER_INDEX_HTML[$($bytes.Length)] PROGMEM = {$bytesText};"
$sizeLine = "const size_t ESPHOME_WEBSERVER_INDEX_HTML_SIZE = $($bytes.Length);"

$mainCpp = [regex]::Replace(
  $mainCpp,
  'const uint8_t ESPHOME_WEBSERVER_INDEX_HTML\[\d+\] PROGMEM = \{.*?\};',
  [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $arrayLine },
  [System.Text.RegularExpressions.RegexOptions]::Singleline
)
$mainCpp = [regex]::Replace(
  $mainCpp,
  'const size_t ESPHOME_WEBSERVER_INDEX_HTML_SIZE = \d+;',
  [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $sizeLine }
)

if ($mainCpp -notmatch 'effective_satellite_name->set_internal\(true\);') {
  $mainCpp = $mainCpp.Replace(
    'effective_satellite_name->set_name("Effective Satellite Name", 3066291396UL);',
    "effective_satellite_name->set_name(`"Effective Satellite Name`", 3066291396UL);`r`n  effective_satellite_name->set_internal(true);"
  )
}

Set-Content -Path $mainCppPath -Value $mainCpp -Encoding UTF8
