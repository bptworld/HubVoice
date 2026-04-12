$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
# Keep setup and runtime on the same per-user state location.
if ($IsWindows) {
  $userDataRoot = if ([string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) { Join-Path $HOME "AppData\Local" } else { $env:LOCALAPPDATA }
  $userDataDir = Join-Path $userDataRoot "HubVoiceSat"
} elseif ($IsMacOS) {
  $userDataDir = Join-Path $HOME "Library/Application Support/HubVoiceSat"
} else {
  $xdgConfigHome = if ([string]::IsNullOrWhiteSpace($env:XDG_CONFIG_HOME)) { Join-Path $HOME ".config" } else { $env:XDG_CONFIG_HOME }
  $userDataDir = Join-Path $xdgConfigHome "hubvoicesat"
}
New-Item -ItemType Directory -Path $userDataDir -Force | Out-Null

$secretsPath = Join-Path $root "secrets.yaml"
$satellitesPath = Join-Path $userDataDir "satellites.csv"
$setupConfigPath = Join-Path $userDataDir "hubvoice-sat-setup.json"
$piperVoicesPath = Join-Path $root "piper_voices"
$yamlPath = Join-Path $root "hubvoice-sat.yaml"
$sonosSpeakersPath = Join-Path $userDataDir "sonos-speakers.csv"
$dlnaSpeakersPath = Join-Path $userDataDir "dlna-speakers.csv"
$setupPort = 8093
$setupSchemaVersion = "2"
try {
  if ($env:HUBVOICESAT_SETUP_PORT) {
    $candidate = [int]$env:HUBVOICESAT_SETUP_PORT
    if ($candidate -ge 1024 -and $candidate -le 65535) {
      $setupPort = $candidate
    }
  }
} catch {
}
$url = "http://127.0.0.1:$setupPort/"
$script:ShutdownRequested = $false
$script:ShutdownAllRequested = $false

function Stop-ProcessById([int]$processId, [string]$reason) {
  if (-not $processId) { return $false }
  try {
    if ($processId -eq $PID) { return $false }
    Stop-Process -Id $processId -Force -ErrorAction Stop
    Write-Host "Stopped PID $processId ($reason)"
    return $true
  } catch {
    return $false
  }
}

function Stop-HubVoiceRuntimeProcesses {
  $stopped = New-Object System.Collections.Generic.HashSet[int]

  # Stop python processes that are running hubvoice-runtime.py from this workspace.
  try {
    $runtimeProcesses = Get-CimInstance Win32_Process -Filter "Name = 'python.exe' OR Name = 'pythonw.exe'"
    foreach ($proc in $runtimeProcesses) {
      $cmd = [string]$proc.CommandLine
      if (-not $cmd) { continue }
      if ($cmd -match "hubvoice-runtime\.py" -and $cmd -like "*$root*") {
        if (Stop-ProcessById ([int]$proc.ProcessId) "runtime script") {
          [void]$stopped.Add([int]$proc.ProcessId)
        }
      }
    }
  } catch {
  }

  # Stop standalone frozen runtime executable from this workspace.
  try {
    $runtimeExeProcesses = Get-CimInstance Win32_Process -Filter "Name = 'HubVoiceRuntime.exe'"
    foreach ($proc in $runtimeExeProcesses) {
      $exePath = [string]$proc.ExecutablePath
      if ($exePath -and $exePath -like "*$root*") {
        if (Stop-ProcessById ([int]$proc.ProcessId) "runtime exe") {
          [void]$stopped.Add([int]$proc.ProcessId)
        }
      }
    }
  } catch {
  }

  # Stop local listener by configured runtime port if bound on this machine.
  try {
    $config = Get-SetupConfig
    $runtimePort = Get-UrlPort $config.hubvoice_url
    if ($runtimePort) {
      $runtimeHost = Get-UrlHost $config.hubvoice_url
      $hostLooksLocal = ($runtimeHost -in @("127.0.0.1", "localhost", "", $env:COMPUTERNAME))
      if ($hostLooksLocal) {
        $connections = Get-NetTCPConnection -State Listen -LocalPort $runtimePort -ErrorAction SilentlyContinue
        foreach ($conn in @($connections)) {
          $pidToStop = [int]$conn.OwningProcess
          if (Stop-ProcessById $pidToStop "runtime port $runtimePort") {
            [void]$stopped.Add($pidToStop)
          }
        }
      }
    }
  } catch {
  }

  return @($stopped)
}

function Stop-LauncherProcesses {
  $stopped = New-Object System.Collections.Generic.HashSet[int]
  try {
    $launcherPath = Get-LauncherPath
    if (-not (Test-Path $launcherPath)) {
      return @($stopped)
    }

    $launcherName = [System.IO.Path]::GetFileName($launcherPath)
    $launcherProcesses = Get-CimInstance Win32_Process -Filter "Name = '$launcherName'"
    foreach ($proc in $launcherProcesses) {
      $exePath = [string]$proc.ExecutablePath
      if ($exePath -and ($exePath -ieq $launcherPath)) {
        if (Stop-ProcessById ([int]$proc.ProcessId) "setup launcher") {
          [void]$stopped.Add([int]$proc.ProcessId)
        }
      }
    }
  } catch {
  }
  return @($stopped)
}

function Stop-SetupWebProcesses {
  $stopped = New-Object System.Collections.Generic.HashSet[int]

  # Stop ALL setup-web.ps1 processes (any path â€” includes AppData-installed copies).
  try {
    $setupProcesses = Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe' OR Name = 'pwsh.exe'"
    foreach ($proc in $setupProcesses) {
      $cmd = [string]$proc.CommandLine
      if (-not $cmd) { continue }
      if ($cmd -match "setup-web\.ps1") {
        if (Stop-ProcessById ([int]$proc.ProcessId) "setup web script") {
          [void]$stopped.Add([int]$proc.ProcessId)
        }
      }
    }
  } catch {
  }

  # Stop any remaining listeners on the setup port window (except this process).
  try {
    for ($port = $setupPort; $port -lt ($setupPort + 10); $port++) {
      $connections = Get-NetTCPConnection -State Listen -LocalPort $port -ErrorAction SilentlyContinue
      foreach ($conn in @($connections)) {
        $pidToStop = [int]$conn.OwningProcess
        if (Stop-ProcessById $pidToStop "setup port $port") {
          [void]$stopped.Add($pidToStop)
        }
      }
    }
  } catch {
  }

  return @($stopped)
}

function Invoke-FullShutdown {
  $setupPids = @(Stop-SetupWebProcesses)
  $runtimePids = @(Stop-HubVoiceRuntimeProcesses)
  $launcherPids = @(Stop-LauncherProcesses)
  $allStopped = @($setupPids + $runtimePids + $launcherPids)
  return @{
    setup_pids = $setupPids
    runtime_pids = $runtimePids
    launcher_pids = $launcherPids
    stopped_count = $allStopped.Count
  }
}

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

function Get-LauncherPath {
  $launcherPath = [string]$env:HUBVOICESAT_LAUNCHER_PATH
  if (-not $launcherPath) {
    $launcherPath = Join-Path $root "HubVoiceSatSetup.exe"
  }
  return $launcherPath
}

function Get-LauncherVersion {
  $launcherPath = Get-LauncherPath
  if (-not (Test-Path $launcherPath)) {
    return "not_found"
  }

  try {
    $fileInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($launcherPath)
    if ($fileInfo -and $fileInfo.FileVersion) {
      return [string]$fileInfo.FileVersion
    }
  } catch {
  }

  return "unknown"
}

function Test-PopulatedSatellitesText([string]$text) {
  if (-not $text) {
    return $false
  }

  foreach ($line in ($text -split "`r?`n")) {
    $row = $line.Trim()
    if (-not $row -or $row.StartsWith("#")) {
      continue
    }
    if ($row.Contains(",")) {
      return $true
    }
  }

  return $false
}

function Get-SecretsState {
  $state = @{
    wifi_ssid = ""
    wifi_password = ""
    wifi_ssid_saved = $false
    wifi_password_saved = $false
  }

  if (-not (Test-Path $secretsPath)) {
    return $state
  }

  foreach ($line in Get-Content $secretsPath) {
    if ($line -match '^\s*wifi_ssid:\s*(.*?)\s*$') {
      $ssidValue = [string]$matches[1]
      if (($ssidValue.StartsWith('"') -and $ssidValue.EndsWith('"')) -or ($ssidValue.StartsWith("'") -and $ssidValue.EndsWith("'"))) {
        if ($ssidValue.Length -ge 2) {
          $ssidValue = $ssidValue.Substring(1, $ssidValue.Length - 2)
        }
      }
      $state.wifi_ssid = $ssidValue
      $state.wifi_ssid_saved = [bool]$ssidValue
    } elseif ($line -match '^\s*wifi_password:\s*(.*?)\s*$') {
      $passwordValue = [string]$matches[1]
      if (($passwordValue.StartsWith('"') -and $passwordValue.EndsWith('"')) -or ($passwordValue.StartsWith("'") -and $passwordValue.EndsWith("'"))) {
        if ($passwordValue.Length -ge 2) {
          $passwordValue = $passwordValue.Substring(1, $passwordValue.Length - 2)
        }
      }
      $state.wifi_password = $passwordValue
      $state.wifi_password_saved = [bool]$passwordValue
    }
  }

  return $state
}

function Save-SecretsState([hashtable]$payload) {
  $existing = Get-SecretsState
  $ssid = [string]$payload.wifi_ssid
  $password = [string]$payload.wifi_password

  # Preserve existing values when the browser submits empty fields.
  if (-not $ssid) {
    $ssid = [string]$existing.wifi_ssid
  }
  if (-not $password) {
    $password = [string]$existing.wifi_password
  }

  $ssid = $ssid.Replace('"', '\"')
  $password = $password.Replace('"', '\"')
  $content = @(
    "wifi_ssid: ""$ssid"""
    "wifi_password: ""$password"""
  )
  Set-Content -Path $secretsPath -Value $content -Encoding UTF8
}

function Get-SatellitesText {
  if (-not (Test-Path $satellitesPath)) {
    return ""
  }
  return ((Get-Content $satellitesPath) -join "`r`n")
}

function Save-SatellitesText([string]$text) {
  $existingAliasByKey = @{}
  $existingAliasByName = @{}
  if (Test-Path $satellitesPath) {
    foreach ($existingLine in (Get-Content $satellitesPath)) {
      $existingRow = $existingLine.Trim()
      if (-not $existingRow -or $existingRow.StartsWith("#")) {
        continue
      }
      $existingParts = $existingRow.Split(",", 3)
      if ($existingParts.Count -lt 2) {
        continue
      }
      $existingName = $existingParts[0].Trim()
      $existingIp = $existingParts[1].Trim()
      $existingAlias = if ($existingParts.Count -ge 3) { $existingParts[2].Trim() } else { "" }
      if ($existingName -and $existingIp -and $existingAlias) {
        $existingKey = ("{0}|{1}" -f $existingName.ToLowerInvariant(), $existingIp.ToLowerInvariant())
        $existingAliasByKey[$existingKey] = $existingAlias
        $existingAliasByName[$existingName.ToLowerInvariant()] = $existingAlias
      }
    }
  }

  $rows = @()
  foreach ($rawLine in ($text -split "`r?`n")) {
    $line = $rawLine.Trim()
    if (-not $line) {
      continue
    }
    if ($line -notmatch ',') {
      throw "Each satellite line must be in the format id,ip[,alias]"
    }
    $parts = $line.Split(",", 3)
    $aliasProvided = $parts.Count -ge 3
    $name = $parts[0].Trim()
    $ip = $parts[1].Trim()
    $alias = ""
    if ($aliasProvided) {
      $alias = $parts[2].Trim()
    }

    # Recover malformed entries like "id,192.168.4.135.livingroom".
    if (-not $aliasProvided -and -not $alias) {
      $ipAliasMatch = [regex]::Match($ip, '^(\d{1,3}(?:\.\d{1,3}){3})\.(.+)$')
      if ($ipAliasMatch.Success) {
        $ip = $ipAliasMatch.Groups[1].Value
        $alias = $ipAliasMatch.Groups[2].Value.Trim()
      }
    }

    if (-not $name -or -not $ip) {
      throw "Each satellite line must include both name and ip"
    }

    # Preserve existing alias when a save only includes id,ip.
    if (-not $aliasProvided -and -not $alias) {
      $key = ("{0}|{1}" -f $name.ToLowerInvariant(), $ip.ToLowerInvariant())
      if ($existingAliasByKey.ContainsKey($key)) {
        $alias = [string]$existingAliasByKey[$key]
      } elseif ($existingAliasByName.ContainsKey($name.ToLowerInvariant())) {
        $alias = [string]$existingAliasByName[$name.ToLowerInvariant()]
      }
    }

    if ($alias) {
      $rows += "$name,$ip,$alias"
    } else {
      $rows += "$name,$ip"
    }
  }

  if ($rows.Count -eq 0) {
    if (Test-Path $satellitesPath) {
      Remove-Item $satellitesPath -Force
    }
  } else {
    Set-Content -Path $satellitesPath -Value $rows -Encoding UTF8
  }
}

function Get-SonosText {
  if (-not (Test-Path $sonosSpeakersPath)) {
    return ""
  }
  return ((Get-Content $sonosSpeakersPath) -join "`r`n")
}

function Save-SonosText([string]$text) {
  $rows = @()
  foreach ($rawLine in ($text -split "`r?`n")) {
    $line = $rawLine.Trim()
    if (-not $line) {
      continue
    }
    if ($line -notmatch ',') {
      throw "Each Sonos speaker line must be in the format name,ip[,alias]"
    }
    $parts = $line.Split(",", 3)
    $name = $parts[0].Trim()
    $ip = $parts[1].Trim()
    $alias = if ($parts.Count -ge 3) { $parts[2].Trim() } else { "" }

    if (-not $name -or -not $ip) {
      throw "Each Sonos speaker line must include both name and ip"
    }

    if ($alias) {
      $rows += "$name,$ip,$alias"
    } else {
      $rows += "$name,$ip"
    }
  }

  if ($rows.Count -eq 0) {
    if (Test-Path $sonosSpeakersPath) {
      Remove-Item $sonosSpeakersPath -Force
    }
  } else {
    Set-Content -Path $sonosSpeakersPath -Value $rows -Encoding UTF8
  }
}

function Get-DlnaText {
  if (-not (Test-Path $dlnaSpeakersPath)) {
    return ""
  }
  return ((Get-Content $dlnaSpeakersPath) -join "`r`n")
}

function Save-DlnaText([string]$text) {
  $rows = @()
  foreach ($rawLine in ($text -split "`r?`n")) {
    $line = $rawLine.Trim()
    if (-not $line) {
      continue
    }
    if ($line -notmatch ',') {
      throw "Each DLNA/UPnP speaker line must be in the format name,ip[,alias]"
    }
    $parts = $line.Split(",", 3)
    $name = $parts[0].Trim()
    $ip = $parts[1].Trim()
    $alias = if ($parts.Count -ge 3) { $parts[2].Trim() } else { "" }

    if (-not $name -or -not $ip) {
      throw "Each DLNA/UPnP speaker line must include both name and ip"
    }

    if ($alias) {
      $rows += "$name,$ip,$alias"
    } else {
      $rows += "$name,$ip"
    }
  }

  if ($rows.Count -eq 0) {
    if (Test-Path $dlnaSpeakersPath) {
      Remove-Item $dlnaSpeakersPath -Force
    }
  } else {
    Set-Content -Path $dlnaSpeakersPath -Value $rows -Encoding UTF8
  }
}

function Convert-SpeakerTextToRows([string]$text) {
  $items = @()
  foreach ($rawLine in ($text -split "`r?`n")) {
    $line = $rawLine.Trim()
    if (-not $line -or $line -notmatch ',') {
      continue
    }
    $parts = $line.Split(",", 3)
    if ($parts.Count -lt 2) {
      continue
    }
    $name = $parts[0].Trim()
    $ip = $parts[1].Trim()
    $alias = if ($parts.Count -ge 3) { $parts[2].Trim() } else { "" }
    if (-not $name -or -not $ip) {
      continue
    }
    $items += @{ name = $name; ip = $ip; alias = $alias }
  }
  return @($items)
}

function Get-DlnaFriendlyNameFromLocation([string]$location) {
  if (-not $location) {
    return $null
  }
  try {
    $response = Invoke-WebRequest -UseBasicParsing -Uri $location -Method GET -TimeoutSec 2
    $content = [string]$response.Content
    if ($content -match '<friendlyName>\s*([^<]+)\s*</friendlyName>') {
      return ([string]$matches[1]).Trim()
    }
  } catch {
  }
  return $null
}

function Discover-DlnaSpeakers([int]$timeoutMs = 2500) {
  $results = New-Object System.Collections.Generic.List[hashtable]
  $seenIps = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  $client = $null

  try {
    $client = New-Object System.Net.Sockets.UdpClient
    $client.EnableBroadcast = $true
    $client.MulticastLoopback = $false
    $client.Client.ReceiveTimeout = 400

    $search = (
      "M-SEARCH * HTTP/1.1`r`n" +
      "HOST: 239.255.255.250:1900`r`n" +
      "MAN: ""ssdp:discover""`r`n" +
      "MX: 2`r`n" +
      "ST: urn:schemas-upnp-org:device:MediaRenderer:1`r`n`r`n"
    )
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($search)
    $endpoint = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Parse("239.255.255.250"), 1900)
    1..2 | ForEach-Object { [void]$client.Send($bytes, $bytes.Length, $endpoint) }

    $deadline = [DateTime]::UtcNow.AddMilliseconds([Math]::Max(500, $timeoutMs))
    while ([DateTime]::UtcNow -lt $deadline) {
      try {
        $remote = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)
        $packet = $client.Receive([ref]$remote)
        if (-not $packet) { continue }
        $text = [System.Text.Encoding]::ASCII.GetString($packet)
        $location = ""
        foreach ($line in ($text -split "`r?`n")) {
          if ($line -match '^(?i)LOCATION:\s*(.+)$') {
            $location = $matches[1].Trim()
            break
          }
        }

        $ip = ""
        try {
          if ($location) {
            $ip = ([Uri]$location).Host
          }
        } catch {
        }
        if (-not $ip) {
          $ip = [string]$remote.Address
        }
        if (-not $ip -or -not $seenIps.Add($ip)) {
          continue
        }

        $friendly = Get-DlnaFriendlyNameFromLocation $location
        if (-not $friendly) {
          $friendly = "DLNA-$($ip -replace '\\.', '-')"
        }
        $id = ($friendly -replace '[^A-Za-z0-9_-]+', '-').Trim('-')
        if (-not $id) {
          $id = "dlna-$($ip -replace '\\.', '-')"
        }

        $results.Add(@{
          name = $id
          ip = $ip
          alias = $friendly
        }) | Out-Null
      } catch [System.Management.Automation.MethodInvocationException] {
        continue
      } catch {
        continue
      }
    }
  } catch {
  } finally {
    if ($client) {
      try { $client.Close() } catch {}
    }
  }

  return @($results)
}

function Add-DiscoveredDlnaSpeakers([hashtable[]]$discovered) {
  $existingRows = @(Convert-SpeakerTextToRows (Get-DlnaText))
  $rowsByIp = @{}
  $rowsByName = @{}
  foreach ($row in $existingRows) {
    $rowsByIp[[string]$row.ip] = $row
    $rowsByName[[string]$row.name] = $row
  }

  $added = 0
  foreach ($item in @($discovered)) {
    $name = [string]$item.name
    $ip = [string]$item.ip
    $alias = [string]$item.alias
    if (-not $name -or -not $ip) { continue }
    if ($rowsByIp.ContainsKey($ip) -or $rowsByName.ContainsKey($name)) { continue }
    $row = @{ name = $name; ip = $ip; alias = $alias }
    $existingRows += $row
    $rowsByIp[$ip] = $row
    $rowsByName[$name] = $row
    $added++
  }

  $lines = @()
  foreach ($row in $existingRows) {
    if ([string]::IsNullOrWhiteSpace([string]$row.alias)) {
      $lines += "{0},{1}" -f $row.name, $row.ip
    } else {
      $lines += "{0},{1},{2}" -f $row.name, $row.ip, $row.alias
    }
  }
  Save-DlnaText ($lines -join "`r`n")
  return $added
}

function Discover-SonosSpeakers([int]$timeoutMs = 2500) {
  $results = New-Object System.Collections.Generic.List[hashtable]
  $seenIps = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  $client = $null

  try {
    $client = New-Object System.Net.Sockets.UdpClient
    $client.EnableBroadcast = $true
    $client.MulticastLoopback = $false
    $client.Client.ReceiveTimeout = 400

    $search = (
      "M-SEARCH * HTTP/1.1`r`n" +
      "HOST: 239.255.255.250:1900`r`n" +
      "MAN: ""ssdp:discover""`r`n" +
      "MX: 2`r`n" +
      "ST: urn:schemas-upnp-org:device:ZonePlayer:1`r`n`r`n"
    )
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($search)
    $endpoint = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Parse("239.255.255.250"), 1900)
    1..2 | ForEach-Object { [void]$client.Send($bytes, $bytes.Length, $endpoint) }

    $deadline = [DateTime]::UtcNow.AddMilliseconds([Math]::Max(500, $timeoutMs))
    while ([DateTime]::UtcNow -lt $deadline) {
      try {
        $remote = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)
        $packet = $client.Receive([ref]$remote)
        if (-not $packet) { continue }
        $text = [System.Text.Encoding]::ASCII.GetString($packet)
        $location = ""
        foreach ($line in ($text -split "`r?`n")) {
          if ($line -match '^(?i)LOCATION:\s*(.+)$') {
            $location = $matches[1].Trim()
            break
          }
        }

        $ip = ""
        try {
          if ($location) {
            $ip = ([Uri]$location).Host
          }
        } catch {
        }
        if (-not $ip) {
          $ip = [string]$remote.Address
        }
        if (-not $ip -or -not $seenIps.Add($ip)) {
          continue
        }

        $friendly = Get-DlnaFriendlyNameFromLocation $location
        if (-not $friendly) {
          $friendly = "Sonos-$($ip -replace '\\.', '-')"
        }
        $id = ($friendly -replace '[^A-Za-z0-9_-]+', '-').Trim('-')
        if (-not $id) {
          $id = "sonos-$($ip -replace '\\.', '-')"
        }

        $results.Add(@{
          name = $id
          ip = $ip
          alias = $friendly
        }) | Out-Null
      } catch [System.Management.Automation.MethodInvocationException] {
        continue
      } catch {
        continue
      }
    }
  } catch {
  } finally {
    if ($client) {
      try { $client.Close() } catch {}
    }
  }

  return @($results)
}

function Add-DiscoveredSonosSpeakers([hashtable[]]$discovered) {
  $existingRows = @(Convert-SpeakerTextToRows (Get-SonosText))
  $rowsByIp = @{}
  $rowsByName = @{}
  foreach ($row in $existingRows) {
    $rowsByIp[[string]$row.ip] = $row
    $rowsByName[[string]$row.name] = $row
  }

  $added = 0
  foreach ($item in @($discovered)) {
    $name = [string]$item.name
    $ip = [string]$item.ip
    $alias = [string]$item.alias
    if (-not $name -or -not $ip) { continue }
    if ($rowsByIp.ContainsKey($ip) -or $rowsByName.ContainsKey($name)) { continue }
    $row = @{ name = $name; ip = $ip; alias = $alias }
    $existingRows += $row
    $rowsByIp[$ip] = $row
    $rowsByName[$name] = $row
    $added++
  }

  $lines = @()
  foreach ($row in $existingRows) {
    if ([string]::IsNullOrWhiteSpace([string]$row.alias)) {
      $lines += "{0},{1}" -f $row.name, $row.ip
    } else {
      $lines += "{0},{1},{2}" -f $row.name, $row.ip, $row.alias
    }
  }
  Save-SonosText ($lines -join "`r`n")
  return $added
}

function Get-SetupConfig {
  $default = @{
    hubvoice_url = ""
    hubitat_host = ""
    hubitat_app_id = ""
    hubitat_access_token = ""
    callback_url = ""
    piper_voice_model = "piper_voices\en_US-amy-medium.onnx"
  }

  if (-not (Test-Path $setupConfigPath)) {
    return $default
  }

  try {
    $loaded = Get-Content $setupConfigPath -Raw | ConvertFrom-Json
    foreach ($key in @("hubvoice_url", "hubitat_host", "hubitat_app_id", "hubitat_access_token", "callback_url", "piper_voice_model")) {
      if ($null -ne $loaded.$key) {
        $default[$key] = [string]$loaded.$key
      }
    }

    return $default
  } catch {
    return $default
  }
}

function Get-PiperVoiceOptions {
  $voices = @()
  if (Test-Path $piperVoicesPath) {
    $voices = Get-ChildItem -Path $piperVoicesPath -Filter *.onnx |
      Sort-Object Name |
      ForEach-Object { "piper_voices\$($_.Name)" }
  }
  return @($voices)
}

function Save-SetupConfig([hashtable]$payload) {
  $existing = Get-SetupConfig
  $accessToken = [string]$payload.hubitat_access_token
  if (-not $accessToken) {
    $accessToken = [string]$existing.hubitat_access_token
  }

  $config = @{
    hubvoice_url = [string]$payload.hubvoice_url
    hubitat_host = [string]$payload.hubitat_host
    hubitat_app_id = [string]$payload.hubitat_app_id
    hubitat_access_token = $accessToken
    callback_url = [string]$payload.callback_url
    piper_voice_model = [string]$payload.piper_voice_model
  }

  $json = $config | ConvertTo-Json -Depth 3
  Set-Content -Path $setupConfigPath -Value $json -Encoding UTF8
}

function Get-State {
  $secrets = Get-SecretsState
  $config = Get-SetupConfig
  return @{
    wifi_ssid = $secrets.wifi_ssid
    wifi_password = ""
    wifi_ssid_saved = [bool]$secrets.wifi_ssid_saved
    wifi_password_saved = [bool]$secrets.wifi_password_saved
    satellites_text = Get-SatellitesText
    sonos_speakers_text = Get-SonosText
    dlna_speakers_text = Get-DlnaText
    hubvoice_url = $config.hubvoice_url
    hubitat_host = $config.hubitat_host
    hubitat_app_id = $config.hubitat_app_id
    hubitat_access_token = ""
    hubitat_access_token_saved = [bool]$config.hubitat_access_token
    callback_url = $config.callback_url
    piper_voice_model = $config.piper_voice_model
    piper_voice_models = @(Get-PiperVoiceOptions)
    launcher_version = Get-LauncherVersion
    launcher_path = Get-LauncherPath
    runtime_state_dir = $userDataDir
    runtime_satellites_path = $satellitesPath
    runtime_sonos_speakers_path = $sonosSpeakersPath
    runtime_dlna_speakers_path = $dlnaSpeakersPath
    runtime_config_path = $setupConfigPath
    setup_schema_version = $setupSchemaVersion
    firmware_target_version = [string](Get-YamlValue -Path $yamlPath -Key "firmware_version")
  }
}

function Get-ProcessSnapshots {
  param(
    [string[]]$Names = @()
  )

  $items = @()

  $onWindows = (($env:OS -eq "Windows_NT") -or ((Get-Variable -Name IsWindows -ErrorAction SilentlyContinue) -and $IsWindows))

  if ($onWindows) {
    try {
      $filter = $null
      if ($Names -and $Names.Count -gt 0) {
        $clauses = @($Names | ForEach-Object { "Name = '$_'" })
        $filter = ($clauses -join ' OR ')
      }

      $processes = if ($filter) {
        Get-CimInstance Win32_Process -Filter $filter -ErrorAction SilentlyContinue
      } else {
        Get-CimInstance Win32_Process -ErrorAction SilentlyContinue
      }

      foreach ($proc in @($processes)) {
        $items += @{
          pid = [int]$proc.ProcessId
          parent_pid = [int]$proc.ParentProcessId
          name = [string]$proc.Name
          path = [string]$proc.ExecutablePath
          command_line = [string]$proc.CommandLine
        }
      }
    } catch {
    }

    return @($items)
  }

  try {
    $output = & ps -eo pid=,ppid=,comm=,args= 2>$null
    foreach ($line in @($output)) {
      $raw = [string]$line
      if (-not $raw) { continue }
      if ($raw -notmatch '^\s*(\d+)\s+(\d+)\s+(\S+)\s*(.*)$') { continue }

      $name = [string]$matches[3]
      if ($Names -and $Names.Count -gt 0 -and ($Names -notcontains $name)) {
        continue
      }

      $items += @{
        pid = [int]$matches[1]
        parent_pid = [int]$matches[2]
        name = $name
        path = ""
        command_line = [string]$matches[4]
      }
    }
  } catch {
  }

  return @($items)
}

function Invoke-RuntimeJsonEndpoint {
  param(
    [string]$BaseUrl,
    [string]$Path,
    [int]$TimeoutSec = 3
  )

  $target = if ($BaseUrl) { $BaseUrl.TrimEnd('/') + $Path } else { "" }
  if (-not $target) {
    return @{ ok = $false; target = $target; error = "HubVoice URL is not configured." }
  }

  try {
    $data = Invoke-RestMethod -Method GET -Uri $target -TimeoutSec $TimeoutSec
    return @{ ok = $true; target = $target; data = $data }
  } catch {
    return @{ ok = $false; target = $target; error = $_.Exception.Message }
  }
}

function Get-SetupWebProcessSnapshot {
  $items = @()
  try {
    $processes = @(Get-ProcessSnapshots -Names @('powershell.exe', 'pwsh.exe', 'pwsh'))
    foreach ($proc in $processes) {
      $cmd = [string]$proc.command_line
      if (-not $cmd) { continue }

      $isSetupWeb = $false
      if ($cmd -match 'setup-web\.ps1') {
        $isSetupWeb = $true
      } elseif ($cmd -match '-File\s+"?([^"\s]+\.ps1)"?') {
        $wrapperPath = $Matches[1]
        if (Test-Path $wrapperPath) {
          try {
            $wrapperText = Get-Content -Path $wrapperPath -Raw -ErrorAction Stop
            if ($wrapperText -match 'setup-web\.ps1' -or $wrapperText -match 'HUBVOICESAT_SETUP_PORT') {
              $isSetupWeb = $true
            }
          } catch {
          }
        }
      }

      if (-not $isSetupWeb) { continue }

      $items += @{
        pid = [int]$proc.pid
        parent_pid = [int]$proc.parent_pid
        name = [string]$proc.name
        command_line = $cmd
      }
    }
  } catch {
  }

  return @($items)
}

function Get-DebugSnapshot {
  $state = Get-State
  $status = Get-StatusSnapshot
  $hubvoiceUrl = [string]$state.hubvoice_url
  $runtimeHealth = Invoke-RuntimeJsonEndpoint -BaseUrl $hubvoiceUrl -Path "/health" -TimeoutSec 3
  $runtimeSatellites = Invoke-RuntimeJsonEndpoint -BaseUrl $hubvoiceUrl -Path "/satellites" -TimeoutSec 3

  $runtimeExe = @()
  $launcherExe = @()

  try {
    $runtimeCandidates = @(Get-ProcessSnapshots -Names @('HubVoiceRuntime.exe', 'python.exe', 'python3', 'python3.exe', 'py.exe'))
    $runtimeExe = @($runtimeCandidates | Where-Object {
      $name = [string]$_.name
      $cmd = [string]$_.command_line
      if ($name -ieq 'HubVoiceRuntime.exe') { return $true }
      return ($cmd -match 'hubvoice-runtime\.py')
    } | ForEach-Object {
      @{
        pid = [int]$_.pid
        parent_pid = [int]$_.parent_pid
        name = [string]$_.name
        path = [string]$_.path
        command_line = [string]$_.command_line
      }
    })
  } catch {
  }

  try {
    $launcherExe = @(Get-ProcessSnapshots -Names @('HubVoiceSat.exe', 'HubVoiceSatSetup.exe') | ForEach-Object {
      @{
        pid = [int]$_.pid
        parent_pid = [int]$_.parent_pid
        name = [string]$_.name
        path = [string]$_.path
      }
    })
  } catch {
  }

  return @{
    generated_at = (Get-Date).ToString("o")
    setup_url = $url
    setup_root = $root
    setup_schema_version = $setupSchemaVersion
    setup_page_path = $setupPagePath
    debug_page_path = $debugPagePath
    state = $state
    status = $status
    runtime_health = $runtimeHealth
    runtime_satellites = $runtimeSatellites
    processes = @{
      setup_web = @(Get-SetupWebProcessSnapshot)
      runtime_exe = $runtimeExe
      launcher_exe = $launcherExe
    }
  }
}

function Sanitize-DiagnosticText([string]$text) {
  if (-not $text) { return "" }

  $sanitized = [string]$text
  # Hide common secret-bearing query string parameters.
  $sanitized = [regex]::Replace($sanitized, '(?i)(access_token=)[^&\s"]+', '$1<redacted>')
  $sanitized = [regex]::Replace($sanitized, '(?i)(token=)[^&\s"]+', '$1<redacted>')

  # Hide simple key/value secret lines that may appear in logs or JSON snippets.
  $sanitized = [regex]::Replace($sanitized, '(?im)("?(wifi_password|hubitat_access_token|access_token|token|api_key|password)"?\s*[:=]\s*")([^"]*)(")', '$1<redacted>$4')
  $sanitized = [regex]::Replace($sanitized, '(?im)("?(wifi_password|hubitat_access_token|access_token|token|api_key|password)"?\s*[:=]\s*)([^\r\n,;\s]+)', '$1<redacted>')

  return $sanitized
}

function Get-LogTail {
  param(
    [string]$Path,
    [int]$MaxLines = 220,
    [int]$MaxChars = 120000
  )

  if (-not $Path) {
    return @{ ok = $false; path = $Path; error = "Path not provided" }
  }

  if (-not (Test-Path -LiteralPath $Path)) {
    return @{ ok = $false; path = $Path; error = "File not found" }
  }

  try {
    $lines = @(Get-Content -LiteralPath $Path -Tail $MaxLines -ErrorAction Stop)
    $text = ($lines -join "`n")
    if ($text.Length -gt $MaxChars) {
      $text = $text.Substring($text.Length - $MaxChars)
    }
    return @{
      ok = $true
      path = $Path
      line_count = $lines.Count
      content = (Sanitize-DiagnosticText $text)
    }
  } catch {
    return @{ ok = $false; path = $Path; error = $_.Exception.Message }
  }
}

function Get-SupportBundle {
  $debug = Get-DebugSnapshot
  $state = $debug.state
  $status = $debug.status
  $runtimeHealth = $debug.runtime_health
  $runtimeSatellites = $debug.runtime_satellites

  $runtimeLog = Join-Path (Join-Path $userDataDir "logs") "hubvoice-runtime.log"
  $runtimeErrLog = Join-Path (Join-Path $userDataDir "logs") "hubvoice-runtime-err.log"
  $legacyRuntimeLog = Join-Path (Join-Path $root "logs") "hubvoice-runtime.log"
  $legacyRuntimeErrLog = Join-Path (Join-Path $root "logs") "hubvoice-runtime-err.log"

  $bundle = @{
    bundle_version = "1"
    generated_at = (Get-Date).ToString("o")
    host = $env:COMPUTERNAME
    setup_url = $url
    setup_root = $root
    runtime_state_dir = $state.runtime_state_dir
    setup_schema_version = $setupSchemaVersion
    setup = @{
      state = $state
      status = $status
      processes = $debug.processes
    }
    runtime = @{
      health = $runtimeHealth
      satellites = $runtimeSatellites
    }
    log_tails = @{
      runtime_log = Get-LogTail -Path $runtimeLog
      runtime_err_log = Get-LogTail -Path $runtimeErrLog
      legacy_runtime_log = Get-LogTail -Path $legacyRuntimeLog
      legacy_runtime_err_log = Get-LogTail -Path $legacyRuntimeErrLog
    }
    notes = @(
      "Secrets are redacted where recognized.",
      "Share this bundle when reporting setup/runtime issues."
    )
  }

  return $bundle
}

function Test-HttpUrl([string]$urlToCheck, [int]$timeoutSec = 1) {
  if (-not $urlToCheck) {
    return @{
      ok = $false
      status = "not_set"
      detail = "Not configured"
    }
  }

  try {
    $response = Invoke-WebRequest -UseBasicParsing -Uri $urlToCheck -Method GET -TimeoutSec $timeoutSec
    return @{
      ok = $true
      status = "online"
      detail = "HTTP $($response.StatusCode)"
    }
  } catch {
    $code = $null
    try { $code = [int]$_.Exception.Response.StatusCode.value__ } catch {}
    if ($code) {
      return @{
        ok = $true
        status = "online"
        detail = "HTTP $code"
      }
    }

    return @{
      ok = $false
      status = "offline"
      detail = $_.Exception.Message
    }
  }
}

function Get-FriendlyHttpStatus([hashtable]$result, [string]$kind) {
  if (-not $result) {
    return $result
  }

  $detail = [string]$result.detail
  if ($result.ok -and $detail -eq "HTTP 404") {
    return @{
      ok = $true
      status = "online"
      detail = if ($kind -eq "hubvoice") { "Reachable. The root path returned HTTP 404, but the voice server is running." } else { "Reachable. The endpoint returned HTTP 404." }
    }
  }

  return $result
}

function Get-HubVoiceRuntimeStatus([string]$hubvoiceUrl, [bool]$portOk, [Nullable[int]]$listenerPid = $null) {
  $listenerText = if ($listenerPid) { " (listener PID $listenerPid)" } else { "" }
  if (-not $hubvoiceUrl) {
    return @{
      ok = $false
      status = "not_set"
      detail = "Not configured"
    }
  }

  if (-not $portOk) {
    return @{
      ok = $false
      status = "offline"
      detail = "Runtime port unreachable"
    }
  }

  try {
    $payload = Invoke-RestMethod -UseBasicParsing -Uri $hubvoiceUrl -Method GET -TimeoutSec 2
    # voice_assistant is a dict keyed by satellite ID; prefer any connected bridge.
    $voiceDict = $payload.voice_assistant
    $voice = $null
    $bridgeCount = 0
    if ($voiceDict) {
      $bridgeCount = @($voiceDict.PSObject.Properties).Count
      foreach ($prop in $voiceDict.PSObject.Properties) {
        $candidate = $prop.Value
        if ($candidate -and [bool]$candidate.connected) {
          $voice = $candidate
          break
        }
      }
      if (-not $voice) {
        $first = @($voiceDict.PSObject.Properties)[0]
        if ($first) { $voice = $first.Value }
      }
    }
    if ($voice) {
      $connected = [bool]$voice.connected
      $statusText = [string]$voice.status
      $runtimeAction = [string]$payload.last_action
      $runtimeTranscript = [string]$payload.last_transcript
      $detailText = if ($connected) {
        if ($runtimeAction -and $runtimeAction -ne "idle") {
          "Connected. $runtimeAction$listenerText"
        } elseif ($runtimeTranscript) {
          "Connected. Last heard: $runtimeTranscript$listenerText"
        } else {
          "Connected and ready$listenerText"
        }
      } elseif ($voice.last_error) {
        "Voice pipeline disconnected. ${statusText}: $($voice.last_error)$listenerText"
      } elseif ($statusText -and $statusText -ne "disconnected") {
        "Voice pipeline disconnected. ${statusText}$listenerText"
      } else {
        "Voice pipeline disconnected. Runtime is reachable but not attached to a satellite API client$listenerText"
      }
      return @{
        ok = $connected
        status = if ($connected) { "online" } else { "offline" }
        detail = $detailText
        hint = if ($connected) {
          ""
        } else {
          "Voice commands will fail with 'No API client connected'. Check satellite IP, ESPHome API port 6054 reachability, and restart runtime."
        }
      }
    }

    return @{
      ok = $false
      status = "offline"
      detail = "Runtime reachable, but no voice bridge is registered (${bridgeCount} satellites tracked)$listenerText"
      hint = "Voice commands will fail with 'No API client connected'. Check satellites.csv entries and ensure runtime can reach satellite port 6054."
    }
  } catch {
    $msg = [string]$_.Exception.Message
    if (-not $msg) { $msg = "Unknown runtime status error" }
    return @{
      ok = $false
      status = "offline"
      detail = "Runtime status check failed: $msg$listenerText"
      hint = "The runtime HTTP endpoint responded unexpectedly. Restart runtime and refresh status."
    }
  }
}

function Get-CallbackHttpStatus([string]$callbackUrl, [bool]$portOk) {
  if (-not $callbackUrl) {
    return @{
      ok = $false
      status = "not_set"
      detail = "Not configured"
    }
  }

  if (-not $portOk) {
    return @{
      ok = $false
      status = "offline"
      detail = "Port unreachable"
    }
  }

  return @{
    ok = $true
    status = "online"
    detail = "Port reachable"
  }
}

function Test-TcpPort([string]$hostName, [int]$port, [int]$timeoutMs = 2000) {
  if (-not $hostName -or -not $port) {
    return $false
  }

  $client = $null
  try {
    $client = New-Object System.Net.Sockets.TcpClient
    $iar = $client.BeginConnect($hostName, $port, $null, $null)
    if (-not $iar.AsyncWaitHandle.WaitOne($timeoutMs, $false)) {
      $client.Close()
      return $false
    }
    $client.EndConnect($iar) | Out-Null
    $client.Close()
    return $true
  } catch {
    try {
      if ($client) { $client.Close() }
    } catch {}
    return $false
  }
}

function Get-UrlHost([string]$rawUrl) {
  try {
    if (-not $rawUrl) { return $null }
    return ([Uri]$rawUrl).Host
  } catch {
    return $null
  }
}

function Get-UrlPort([string]$rawUrl) {
  try {
    if (-not $rawUrl) { return $null }
    return ([Uri]$rawUrl).Port
  } catch {
    return $null
  }
}

function Is-HostLocal([string]$hostName) {
  if (-not $hostName) {
    return $false
  }

  $normalized = $hostName.Trim().ToLowerInvariant()
  if ($normalized -in @("127.0.0.1", "localhost", "::1", $env:COMPUTERNAME.ToLowerInvariant())) {
    return $true
  }

  try {
    $localAddresses = New-Object System.Collections.Generic.HashSet[string]
    foreach ($nic in [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()) {
      if ($nic.OperationalStatus -ne [System.Net.NetworkInformation.OperationalStatus]::Up) {
        continue
      }
      foreach ($uni in $nic.GetIPProperties().UnicastAddresses) {
        if ($uni -and $uni.Address) {
          [void]$localAddresses.Add($uni.Address.ToString())
        }
      }
    }

    foreach ($addr in [System.Net.Dns]::GetHostAddresses($hostName)) {
      if ($addr -and $localAddresses.Contains($addr.ToString())) {
        return $true
      }
    }
  } catch {
  }

  return $false
}

function Get-LocalListenerPid([string]$rawUrl) {
  $hostName = Get-UrlHost $rawUrl
  $port = Get-UrlPort $rawUrl
  if (-not $hostName -or -not $port) {
    return $null
  }

  $hostLooksLocal = Is-HostLocal $hostName
  if (-not $hostLooksLocal) {
    return $null
  }

  try {
    $listener = @(Get-NetTCPConnection -State Listen -LocalPort $port -ErrorAction SilentlyContinue | Select-Object -First 1)
    if ($listener.Count -gt 0 -and $listener[0].OwningProcess) {
      return [int]$listener[0].OwningProcess
    }
  } catch {
  }

  return $null
}

function Get-SatelliteFirmwareVersion([string]$ip, [int]$port) {
  if (-not $ip -or -not $port) {
    return $null
  }

  try {
    $targetUrl = "http://$ip`:$port/"
    $response = Invoke-WebRequest -UseBasicParsing -Uri $targetUrl -Method GET -TimeoutSec 2
    $content = [string]$response.Content
    if ($content -match 'Version\s*:\s*([0-9]+\.[0-9]+\.[0-9]+\.[0-9]+)') {
      return $matches[1]
    }
  } catch {
  }

  return $null
}

function Normalize-Version([string]$version) {
  if (-not $version) {
    return $null
  }

  if ($version -notmatch '^\d+\.\d+\.\d+\.\d+$') {
    return $null
  }

  return ($version -split '\.') | ForEach-Object { [int]$_ }
}

function Compare-Version([string]$left, [string]$right) {
  $leftParts = Normalize-Version $left
  $rightParts = Normalize-Version $right
  if (-not $leftParts -or -not $rightParts) {
    return $null
  }

  for ($i = 0; $i -lt 4; $i++) {
    if ($leftParts[$i] -lt $rightParts[$i]) { return -1 }
    if ($leftParts[$i] -gt $rightParts[$i]) { return 1 }
  }

  return 0
}

function Get-SatelliteStatus {
  $items = @()
  # Default firmware target from main YAML.
  $defaultFirmware = [string](Get-YamlValue -Path $yamlPath -Key "firmware_version")
  $csvText = Get-SatellitesText
  foreach ($line in ($csvText -split "`r?`n")) {
    $row = $line.Trim()
    if (-not $row) { continue }

    $parts = $row.Split(",", 3)
    if ($parts.Count -lt 2) { continue }

    $name = $parts[0].Trim()
    $ip = $parts[1].Trim()
    $alias = if ($parts.Count -ge 3) { $parts[2].Trim() } else { "" }
    if (-not $name -or -not $ip) { continue }

    $web8080 = Test-TcpPort $ip 8080 800
    $web80 = if ($web8080) { $false } else { Test-TcpPort $ip 80 800 }
    $webOk = $web8080 -or $web80
    $webPort = if ($web8080) { 8080 } elseif ($web80) { 80 } else { $null }
    $pingOk = $webOk
    $satFirmware = if ($webOk -and $webPort) { Get-SatelliteFirmwareVersion -ip $ip -port $webPort } else { $null }
    $targetFirmware = $defaultFirmware
    $comparison = Compare-Version $satFirmware $targetFirmware
    $updateStatus = if (-not $satFirmware) {
      "unknown"
    } elseif ($comparison -eq 0) {
      "current"
    } elseif ($comparison -lt 0) {
      "outdated"
    } else {
      "ahead"
    }

    $items += @{
      name = $name
      ip = $ip
      alias = $alias
      ping = $pingOk
      web = $webOk
      web_port = $webPort
      firmware_version = if ($satFirmware) { $satFirmware } else { "unknown" }
      target_firmware_version = $targetFirmware
      update_status = $updateStatus
    }
  }

  return @($items)
}

function Get-StatusSnapshot {
  $config = Get-SetupConfig
  $hubitatHealthUrl = ""
  if ($config.hubitat_host -and $config.hubitat_app_id -and $config.hubitat_access_token) {
    $hubitatHealthUrl = "$($config.hubitat_host.TrimEnd('/'))/apps/api/$($config.hubitat_app_id)/health?access_token=$($config.hubitat_access_token)"
  }

  $hubvoiceHost = Get-UrlHost $config.hubvoice_url
  $hubvoicePort = Get-UrlPort $config.hubvoice_url
  $runtimeListenerPid = Get-LocalListenerPid $config.hubvoice_url
  $callbackHost = Get-UrlHost $config.callback_url
  $callbackPort = Get-UrlPort $config.callback_url
  $hubvoicePortOk = [bool](Test-TcpPort $hubvoiceHost $hubvoicePort 800)
  $callbackPortOk = [bool](Test-TcpPort $callbackHost $callbackPort 800)
  $hubvoiceHttp = if ($hubvoicePortOk) { Get-FriendlyHttpStatus (Test-HttpUrl $config.hubvoice_url 1) "hubvoice" } else { @{ ok = $false; status = if ($config.hubvoice_url) { "offline" } else { "not_set" }; detail = if ($config.hubvoice_url) { "Port unreachable" } else { "Not configured" } } }
  $hubvoiceRuntime = Get-HubVoiceRuntimeStatus $config.hubvoice_url $hubvoicePortOk $runtimeListenerPid
  $callbackHttp = Get-CallbackHttpStatus $config.callback_url $callbackPortOk
  $hubitatHttp = if ($hubitatHealthUrl) { Test-HttpUrl $hubitatHealthUrl 1 } else { @{ ok = $false; status = "not_set"; detail = "Not configured" } }

  return @{
    urls = @{
      setup_page = $url
      hubvoice = $config.hubvoice_url
      callback = $config.callback_url
    }
    setup_page = @{
      ok = $true
      status = "online"
      detail = "Listening on $url"
    }
    hubvoice = $hubvoiceHttp
    hubvoice_runtime = $hubvoiceRuntime
    hubvoice_port = @{
      ok = $hubvoicePortOk
      status = if (-not $config.hubvoice_url) { "not_set" } elseif ($hubvoicePortOk) { "online" } else { "offline" }
      detail = if ($hubvoiceHost -and $hubvoicePort) {
        if ($runtimeListenerPid) { "$hubvoiceHost`:$hubvoicePort (listener PID $runtimeListenerPid)" } else { "$hubvoiceHost`:$hubvoicePort" }
      } else { "Not configured" }
    }
    callback = $callbackHttp
    callback_port = @{
      ok = $callbackPortOk
      status = if (-not $config.callback_url) { "not_set" } elseif ($callbackPortOk) { "online" } else { "offline" }
      detail = if ($callbackHost -and $callbackPort) { "$callbackHost`:$callbackPort" } else { "Not configured" }
    }
    hubitat_health = $hubitatHttp
    satellites = @(Get-SatelliteStatus)
    sonos_speakers = @(Convert-SpeakerTextToRows (Get-SonosText) | ForEach-Object {
      $speakerIp = [string]$_.ip
      @{
        name = [string]$_.name
        ip = $speakerIp
        alias = [string]$_.alias
        ping = [bool](Test-TcpPort $speakerIp 1400 500)
        web_port = 1400
      }
    })
    dlna_speakers = @(Convert-SpeakerTextToRows (Get-DlnaText) | ForEach-Object {
      $speakerIp = [string]$_.ip
      @{
        name = [string]$_.name
        ip = $speakerIp
        alias = [string]$_.alias
        ping = [bool](Test-TcpPort $speakerIp 80 500)
        web_port = 80
      }
    })
  }
}

function New-DesktopShortcut {
  $desktop = [Environment]::GetFolderPath("Desktop")
  $shortcutPath = Join-Path $desktop "HubVoiceSat Setup.lnk"
  $scriptPath = Join-Path $root "setup-web.ps1"
  $launcherPath = Get-LauncherPath
  $shell = New-Object -ComObject WScript.Shell
  $shortcut = $shell.CreateShortcut($shortcutPath)

  if (Test-Path $scriptPath) {
    # Prefer running the local script so the shortcut always uses latest workspace updates.
    $shortcut.TargetPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    $quotedScriptPath = '"' + $scriptPath + '"'
    $shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -File $quotedScriptPath"
    $shortcut.WorkingDirectory = $root
  } else {
    $shortcut.TargetPath = $launcherPath
    $shortcut.WorkingDirectory = Split-Path -Parent $launcherPath
  }

  $shortcut.IconLocation = "$env:SystemRoot\System32\Speech\SpeechUX\SpeechUXWiz.exe,0"
  $shortcut.Description = "Open HubVoiceSat setup page"
  $shortcut.Save()
  return $shortcutPath
}

function Open-ReleasesFolder {
  $releasesPath = Join-Path $root "releases"
  New-Item -ItemType Directory -Force -Path $releasesPath | Out-Null
  Start-Process explorer.exe $releasesPath | Out-Null
  return $releasesPath
}

function Open-ExternalUrl([string]$targetUrl) {
  if (-not $targetUrl) {
    throw "URL is required."
  }
  Start-Process $targetUrl | Out-Null
}

function Write-JsonResponse($context, [int]$statusCode, $payload) {
  try {
    $json = $payload | ConvertTo-Json -Depth 5
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $context.Response.StatusCode = $statusCode
    $context.Response.ContentType = "application/json; charset=utf-8"
    $context.Response.ContentEncoding = [System.Text.Encoding]::UTF8
    try { $context.Response.ContentLength64 = $bytes.Length } catch {}
    $context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
  } catch {
  } finally {
    try { $context.Response.OutputStream.Close() } catch {}
  }
}

function Write-TextResponse($context, [int]$statusCode, [string]$contentType, [string]$text) {
  try {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($text)
    $context.Response.StatusCode = $statusCode
    $context.Response.ContentType = $contentType
    $context.Response.ContentEncoding = [System.Text.Encoding]::UTF8
    try { $context.Response.Headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0' } catch {}
    try { $context.Response.Headers['Pragma'] = 'no-cache' } catch {}
    try { $context.Response.Headers['Expires'] = '0' } catch {}
    try { $context.Response.ContentLength64 = $bytes.Length } catch {}
    $context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
  } catch {
  } finally {
    try { $context.Response.OutputStream.Close() } catch {}
  }
}

$setupPagePath = Join-Path $root "_live_setup_page.html"
if (-not (Test-Path $setupPagePath)) {
  throw "Setup page source not found at $setupPagePath"
}

$speakersPagePath = Join-Path $root "_live_speakers_page.html"
if (-not (Test-Path $speakersPagePath)) {
  throw "Speakers page source not found at $speakersPagePath"
}

$debugPagePath = Join-Path $root "_live_setup_debug_page.html"
if (-not (Test-Path $debugPagePath)) {
  throw "Debug page source not found at $debugPagePath"
}

$html = Get-Content -Path $setupPagePath -Raw -Encoding UTF8
$setupUiToken = (Get-Date).ToString("yyyyMMddHHmmss")

function Start-RuntimeIfNeeded {
  $runtimeVenvPythonw = Join-Path $root ".envs\runtime\Scripts\pythonw.exe"
  $runtimeVenvPython = Join-Path $root ".envs\runtime\Scripts\python.exe"
  $venvPythonw = Join-Path $root ".envs\main\Scripts\pythonw.exe"
  $venvPython = Join-Path $root ".envs\main\Scripts\python.exe"
  $runtimeScript = Join-Path $root "hubvoice-runtime.py"
  $pythonExe = $null
  
  if (-not (Test-Path $runtimeScript)) {
    return
  }

  if (Test-Path $runtimeVenvPythonw) {
    $pythonExe = $runtimeVenvPythonw
  } elseif (Test-Path $runtimeVenvPython) {
    $pythonExe = $runtimeVenvPython
  } elseif (Test-Path $venvPythonw) {
    $pythonExe = $venvPythonw
  } elseif (Test-Path $venvPython) {
    $pythonExe = $venvPython
  } else {
    $pythonExe = "py"
  }

  # Prefer the listener PID as source of truth (venv launchers can show an extra shim process).
  $config = Get-SetupConfig
  $runtimeListenerPid = Get-LocalListenerPid $config.hubvoice_url
  if ($runtimeListenerPid) {
    return
  }

  # Fallback process check when no local listener can be resolved.
  $existing = @(Get-CimInstance Win32_Process -Filter "Name = 'python.exe' OR Name = 'pythonw.exe'" -ErrorAction SilentlyContinue | Where-Object {
    $cmd = [string]$_.CommandLine
    $cmd -match "hubvoice-runtime\.py" -and $cmd -like "*$root*"
  })

  if ($existing.Count -gt 0) {
    return
  }

  # Start runtime in background
  try {
    $runtimeArgs = @("-u", $runtimeScript)
    if ($pythonExe -ieq "py") {
      $runtimeArgs = @("-3", "-u", $runtimeScript)
    }

    $process = Start-Process -FilePath $pythonExe `
      -ArgumentList $runtimeArgs `
      -WorkingDirectory $root `
      -WindowStyle Hidden `
      -PassThru `
      -ErrorAction SilentlyContinue
    if ($process) {
      Start-Sleep -Milliseconds 300
    }
  } catch {
  }
}

$listener = $null
$bound = $false

# Ensure runtime is running before starting the web server
Start-RuntimeIfNeeded

# Never reuse existing setup page instances; stop them so latest UI is always served.
try {
  [void](Stop-SetupWebProcesses)
} catch {
}

for ($offset = 0; $offset -lt 10 -and -not $bound; $offset++) {
  $candidatePort = $setupPort + $offset
  $candidateUrl = "http://127.0.0.1:$candidatePort/"
  $candidate = [System.Net.HttpListener]::new()
  $candidate.Prefixes.Add($candidateUrl)

  try {
    $candidate.Start()
    $listener = $candidate
    $url = $candidateUrl
    $bound = $true
  } catch {
    try { $candidate.Close() } catch {}
  }
}

if (-not $bound -or -not $listener) {
  throw "Unable to start setup page listener on ports $setupPort-$($setupPort + 9)."
}

Write-Host ""
Write-Host "HubVoiceSat setup page running at $url"
Write-Host "Press Ctrl+C to stop."
Write-Host ""

try {
  if (-not $env:HUBVOICESAT_SUPPRESS_AUTO_OPEN) {
    Start-Process $url | Out-Null
  }
} catch {
}

while ($listener.IsListening) {
  $context = $listener.GetContext()
  try {
    $path = $context.Request.Url.AbsolutePath
    $method = $context.Request.HttpMethod

    if ($path -eq "/") {
      try {
        $html = Get-Content -Path $setupPagePath -Raw -Encoding UTF8
      } catch {
      }
      Write-TextResponse $context 200 "text/html; charset=utf-8" $html
      continue
    }

    if ($path -eq "/speakers") {
      try {
        $speakersHtml = Get-Content -Path $speakersPagePath -Raw -Encoding UTF8
      } catch {
        $speakersHtml = ""
      }
      Write-TextResponse $context 200 "text/html; charset=utf-8" $speakersHtml
      continue
    }

    if ($path -eq "/debug") {
      try {
        $debugHtml = Get-Content -Path $debugPagePath -Raw -Encoding UTF8
      } catch {
        $debugHtml = ""
      }
      Write-TextResponse $context 200 "text/html; charset=utf-8" $debugHtml
      continue
    }

    if ($path -eq "/favicon.ico") {
      $context.Response.StatusCode = 204
      $context.Response.OutputStream.Close()
      continue
    }

    if ($path -eq "/api/state" -and $method -eq "GET") {
      Write-JsonResponse $context 200 (Get-State)
      continue
    }

    if ($path -eq "/api/version" -and $method -eq "GET") {
      $setupHash = ""
      try {
        $raw = Get-Content -Path $setupPagePath -Raw -Encoding UTF8
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($raw)
        $sha = [System.Security.Cryptography.SHA256]::Create()
        $setupHash = [System.BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-', '')
      } catch {
      }

      Write-JsonResponse $context 200 @{
        ok = $true
        token = $setupUiToken
        root = $root
        setup_page_path = $setupPagePath
        setup_page_hash = $setupHash
      }
      continue
    }

    if ($path -eq "/api/status" -and $method -eq "GET") {
      Write-JsonResponse $context 200 (Get-StatusSnapshot)
      continue
    }

    if ($path -eq "/api/debug_snapshot" -and $method -eq "GET") {
      Write-JsonResponse $context 200 (Get-DebugSnapshot)
      continue
    }

    if ($path -eq "/api/support_bundle" -and $method -eq "GET") {
      Write-JsonResponse $context 200 (Get-SupportBundle)
      continue
    }

    if ($path -eq "/api/speakers/state" -and $method -eq "GET") {
      Write-JsonResponse $context 200 @{
        ok = $true
        sonos_speakers_text = Get-SonosText
        dlna_speakers_text = Get-DlnaText
      }
      continue
    }

    if ($path -eq "/api/save" -and $method -eq "POST") {
      $reader = New-Object System.IO.StreamReader($context.Request.InputStream, $context.Request.ContentEncoding)
      $body = $reader.ReadToEnd()
      $reader.Dispose()
      $payloadObject = $body | ConvertFrom-Json
      $payload = @{
        wifi_ssid = [string]$payloadObject.wifi_ssid
        wifi_password = [string]$payloadObject.wifi_password
        hubvoice_url = [string]$payloadObject.hubvoice_url
        hubitat_host = [string]$payloadObject.hubitat_host
        hubitat_app_id = [string]$payloadObject.hubitat_app_id
        hubitat_access_token = [string]$payloadObject.hubitat_access_token
        callback_url = [string]$payloadObject.callback_url
        piper_voice_model = [string]$payloadObject.piper_voice_model
        satellites_text = [string]$payloadObject.satellites_text
      }

      Save-SecretsState $payload
      Save-SatellitesText ([string]$payload.satellites_text)
      Save-SetupConfig $payload

      Write-JsonResponse $context 200 @{
        ok = $true
        message = "Saved HubVoiceSat setup values."
      }
      continue
    }

    if ($path -eq "/api/create_desktop_shortcut" -and $method -eq "POST") {
      $shortcutPath = New-DesktopShortcut
      Write-JsonResponse $context 200 @{
        ok = $true
        message = "Desktop shortcut created at $shortcutPath"
      }
      continue
    }

    if ($path -eq "/api/build_ota_release" -and $method -eq "POST") {
      $scriptPath = Join-Path $root "build-ota-release.ps1"
      $output = & powershell -NoProfile -ExecutionPolicy Bypass -File $scriptPath 2>&1 | Out-String
      Write-JsonResponse $context 200 @{
        ok = $true
        message = ($output.Trim() -replace "\s+", " ")
      }
      continue
    }

    if ($path -eq "/api/open_releases_folder" -and $method -eq "POST") {
      $releasePath = Open-ReleasesFolder
      Write-JsonResponse $context 200 @{
        ok = $true
        message = "Opened $releasePath"
      }
      continue
    }

    if ($path -eq "/api/speakers/save" -and $method -eq "POST") {
      $reader = New-Object System.IO.StreamReader($context.Request.InputStream, $context.Request.ContentEncoding)
      $body = $reader.ReadToEnd()
      $reader.Dispose()
      $payloadObject = $body | ConvertFrom-Json
      Save-SonosText ([string]$payloadObject.sonos_speakers_text)
      Save-DlnaText ([string]$payloadObject.dlna_speakers_text)
      Write-JsonResponse $context 200 @{
        ok = $true
        message = "Saved speaker targets."
      }
      continue
    }

    if ($path -eq "/api/speakers/discover_dlna" -and $method -eq "POST") {
      $discovered = @(Discover-DlnaSpeakers)
      $added = Add-DiscoveredDlnaSpeakers $discovered
      Write-JsonResponse $context 200 @{
        ok = $true
        message = if ($added -gt 0) { "Discovered $added DLNA speaker(s)." } else { "No new DLNA speakers found." }
        dlna_speakers_text = Get-DlnaText
        discovered = $discovered
      }
      continue
    }

    if ($path -eq "/api/speakers/discover_sonos" -and $method -eq "POST") {
      $discovered = @(Discover-SonosSpeakers)
      $added = Add-DiscoveredSonosSpeakers $discovered
      Write-JsonResponse $context 200 @{
        ok = $true
        message = if ($added -gt 0) { "Discovered $added Sonos speaker(s)." } else { "No new Sonos speakers found." }
        sonos_speakers_text = Get-SonosText
        discovered = $discovered
      }
      continue
    }

    if ($path -eq "/api/shutdown" -and $method -eq "POST") {
      $script:ShutdownRequested = $true
      Write-JsonResponse $context 200 @{
        ok = $true
        message = "Setup server is shutting down."
      }
      continue
    }

    if ($path -eq "/api/shutdown_all" -and $method -eq "POST") {
      $shutdown = Invoke-FullShutdown
      $script:ShutdownRequested = $true
      $script:ShutdownAllRequested = $true
      $count = [int]$shutdown.stopped_count
      $message = if ($count -gt 0) {
        "Shutting down everything. Stopped $count process(es)."
      } else {
        "Shutting down setup server. Runtime/launcher processes were not found."
      }
      Write-JsonResponse $context 200 @{
        ok = $true
        message = $message
        stopped = $shutdown
      }
      continue
    }

    if ($path -eq "/api/open_url" -and $method -eq "POST") {
      $reader = New-Object System.IO.StreamReader($context.Request.InputStream, $context.Request.ContentEncoding)
      $body = $reader.ReadToEnd()
      $reader.Dispose()
      $payloadObject = $body | ConvertFrom-Json
      $targetUrl = [string]$payloadObject.url
      Open-ExternalUrl $targetUrl
      Write-JsonResponse $context 200 @{
        ok = $true
        message = "Opened $targetUrl"
      }
      continue
    }

    Write-JsonResponse $context 404 @{
      ok = $false
      message = "Not found."
    }
  } catch {
    try {
      Write-JsonResponse $context 500 @{
        ok = $false
        message = $_.Exception.Message
      }
    } catch {
    }
  }

  if ($script:ShutdownRequested) {
    try {
      $listener.Stop()
      $listener.Close()
    } catch {
    }
    break
  }
}
