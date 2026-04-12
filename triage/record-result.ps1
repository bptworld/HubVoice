param(
  [Parameter(Mandatory = $true)]
  [string]$Device,

  [Parameter(Mandatory = $true)]
  [ValidateSet('good','bad')]
  [string]$Outcome,

  [Parameter(Mandatory = $true)]
  [string]$WakeWord,

  [Parameter(Mandatory = $true)]
  [string]$I2S,

  [string]$Notes = ""
)

$commit = (git rev-parse --short HEAD).Trim()
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$safeNotes = $Notes -replace '"', '""'

$line = '"{0}","{1}","{2}","{3}","{4}","{5}","{6}"' -f $timestamp, $commit, $Device, $Outcome, $WakeWord, $I2S, $safeNotes
Add-Content -Path "triage/results.csv" -Value $line
Write-Output "Logged: $line"
