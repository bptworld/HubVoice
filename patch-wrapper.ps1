param(
  [Parameter(ValueFromRemainingArguments = $true)]
  [string[]]$RemainingArgs
)

$localPatchExe = Join-Path $PSScriptRoot 'patch.exe'
$localPatchDll = Join-Path $PSScriptRoot 'patch.dll'

if (Test-Path $localPatchExe) {
  & $localPatchExe @RemainingArgs
  exit $LASTEXITCODE
}

if (Test-Path $localPatchDll) {
  & dotnet $localPatchDll @RemainingArgs
  exit $LASTEXITCODE
}

Write-Error "Local patch shim not found at $localPatchExe or $localPatchDll"
exit 1
