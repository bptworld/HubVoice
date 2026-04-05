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

function Get-LegacyStateRoots {
  $roots = @($root)

  $launcherPath = Get-LauncherPath
  if ($launcherPath) {
    $launcherDir = Split-Path -Parent $launcherPath
    if ($launcherDir) {
      $roots += $launcherDir
    }
  }

  try {
    $cwd = (Get-Location).Path
    if ($cwd) {
      $roots += $cwd
    }
  } catch {
  }

  $seen = @{}
  $unique = @()
  foreach ($candidate in $roots) {
    if (-not $candidate) {
      continue
    }
    try {
      $resolved = [System.IO.Path]::GetFullPath($candidate)
    } catch {
      continue
    }
    $key = $resolved.ToLowerInvariant()
    if ($seen.ContainsKey($key)) {
      continue
    }
    $seen[$key] = $true
    $unique += $resolved
  }

  return @($unique)
}

function Migrate-LegacyUserFiles {
  if (-not (Test-Path $satellitesPath)) {
    foreach ($legacyRoot in (Get-LegacyStateRoots)) {
      $legacySatellitesPath = Join-Path $legacyRoot "satellites.csv"
      if (-not (Test-Path $legacySatellitesPath)) {
        continue
      }
      try {
        $legacyText = Get-Content $legacySatellitesPath -Raw -Encoding UTF8
        if (Test-PopulatedSatellitesText $legacyText) {
          Set-Content -Path $satellitesPath -Value $legacyText -Encoding UTF8
          break
        }
      } catch {
      }
    }
  }

  if (-not (Test-Path $setupConfigPath)) {
    foreach ($legacyRoot in (Get-LegacyStateRoots)) {
      $legacyConfigPath = Join-Path $legacyRoot "hubvoice-sat-setup.json"
      if (-not (Test-Path $legacyConfigPath)) {
        continue
      }
      try {
        $legacyRaw = Get-Content $legacyConfigPath -Raw -Encoding UTF8
        $legacyConfig = $legacyRaw | ConvertFrom-Json
        if ($legacyConfig) {
          Set-Content -Path $setupConfigPath -Value $legacyRaw -Encoding UTF8
          break
        }
      } catch {
      }
    }
  }
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
  Migrate-LegacyUserFiles
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

function Get-SetupConfig {
  Migrate-LegacyUserFiles
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
    hubvoice_url = $config.hubvoice_url
    hubitat_host = $config.hubitat_host
    hubitat_app_id = $config.hubitat_app_id
    hubitat_access_token = ""
    hubitat_access_token_saved = [bool]$config.hubitat_access_token
    callback_url = $config.callback_url
    piper_voice_model = $config.piper_voice_model
    piper_voice_models = @(Get-PiperVoiceOptions)
    launcher_version = Get-LauncherVersion
    setup_schema_version = $setupSchemaVersion
    firmware_target_version = [string](Get-YamlValue -Path $yamlPath -Key "firmware_version")
  }
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
  # Default firmware target from main YAML; echos3r satellites get their own target version.
  $defaultFirmware = [string](Get-YamlValue -Path $yamlPath -Key "firmware_version")
  $echos3rYamlPath = Join-Path $root "hubvoice-sat-echos3r.yaml"
  $echos3rFirmware = if (Test-Path $echos3rYamlPath) { [string](Get-YamlValue -Path $echos3rYamlPath -Key "firmware_version") } else { $defaultFirmware }
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
    $targetFirmware = if ($name -like "*atom*" -or $name -like "*echos3r*") { $echos3rFirmware } else { $defaultFirmware }
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
