param(
  [Parameter(ValueFromRemainingArguments = $true)]
  [string[]]$RemainingArgs
)

$localPatchExe = Join-Path $PSScriptRoot 'patch.exe'
$localPatchDll = Join-Path $PSScriptRoot 'patch.dll'
$patchProject = Join-Path $PSScriptRoot 'patch-shim\PatchShim.csproj'

$stripCount = 0
$patchFile = $null
for ($i = 0; $i -lt $RemainingArgs.Count; $i++) {
  $arg = $RemainingArgs[$i]
  if ($arg -eq '-p' -and ($i + 1) -lt $RemainingArgs.Count) {
    $i++
    $stripCount = [int]$RemainingArgs[$i]
    continue
  }
  if ($arg.StartsWith('-p') -and $arg.Length -gt 2) {
    $stripCount = [int]$arg.Substring(2)
    continue
  }
  if ($arg -eq '-i' -and ($i + 1) -lt $RemainingArgs.Count) {
    $i++
    $patchFile = $RemainingArgs[$i]
    continue
  }
  if ($arg.StartsWith('-i') -and $arg.Length -gt 2) {
    $patchFile = $arg.Substring(2)
    continue
  }
}

$git = Get-Command git -ErrorAction SilentlyContinue
if ($git -and $patchFile) {
  & git apply --no-index --binary --whitespace=nowarn ("-p" + $stripCount) $patchFile
  if ($LASTEXITCODE -eq 0) {
    exit 0
  }
}

if (Test-Path $localPatchExe) {
  & $localPatchExe @RemainingArgs
  exit $LASTEXITCODE
}

if (Test-Path $localPatchDll) {
  & dotnet $localPatchDll @RemainingArgs
  exit $LASTEXITCODE
}

if (Test-Path $patchProject) {
  & dotnet run --project $patchProject -- @RemainingArgs
  exit $LASTEXITCODE
}

Write-Error "Unable to apply patch. git apply failed and no local patch shim was available."
exit 1
