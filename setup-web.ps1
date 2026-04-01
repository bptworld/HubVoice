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

  # Stop ALL setup-web.ps1 processes (any path — includes AppData-installed copies).
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

function Get-SecretsState {
  $state = @{
    wifi_ssid = ""
    wifi_password = ""
  }

  if (-not (Test-Path $secretsPath)) {
    return $state
  }

  foreach ($line in Get-Content $secretsPath) {
    if ($line -match '^\s*wifi_ssid:\s*"(.*)"\s*$') {
      $state.wifi_ssid = $matches[1]
    } elseif ($line -match '^\s*wifi_password:\s*"(.*)"\s*$') {
      $state.wifi_password = $matches[1]
    }
  }

  return $state
}

function Save-SecretsState([hashtable]$payload) {
  $ssid = [string]$payload.wifi_ssid
  $password = [string]$payload.wifi_password
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
    wifi_password = $secrets.wifi_password
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
    # voice_assistant is a dict keyed by satellite ID; get the first (or any connected) bridge
    $voiceDict = $payload.voice_assistant
    $voice = $null
    if ($voiceDict) {
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
        "Disconnected. ${statusText}: $($voice.last_error)$listenerText"
      } elseif ($statusText -and $statusText -ne "disconnected") {
        "${statusText}$listenerText"
      } else {
        "Disconnected$listenerText"
      }
      return @{
        ok = $connected
        status = if ($connected) { "online" } else { "offline" }
        detail = $detailText
      }
    }
  } catch {
  }

  return @{
    ok = $true
    status = "online"
    detail = "Runtime reachable$listenerText"
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
    try { $context.Response.ContentLength64 = $bytes.Length } catch {}
    $context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
  } catch {
  } finally {
    try { $context.Response.OutputStream.Close() } catch {}
  }
}

$html = @'
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>HubVoiceSat Setup</title>
  <style>
    body { font-family: Segoe UI, Arial, sans-serif; background:#111; color:#f5f5f5; margin:0; }
    .wrap { max-width: 920px; margin: 0 auto; padding: 24px; }
    .card { background:#1c1c1c; border:1px solid #343434; border-radius:12px; padding:18px; margin-top:16px; }
    h1, h2 { margin:0 0 10px; }
    p { color:#c9c9c9; line-height:1.45; }
    .grid { display:grid; grid-template-columns: repeat(2, 1fr); gap:14px; }
    .full { grid-column: 1 / -1; }
    label { display:block; margin-bottom:6px; color:#ddd; }
    input, textarea { width:100%; box-sizing:border-box; background:#101010; color:#f5f5f5; border:1px solid #404040; border-radius:8px; padding:10px 12px; }
    textarea { min-height: 110px; resize: vertical; font-family: Consolas, monospace; }
    button { background:#2b72d6; color:#fff; border:none; border-radius:8px; padding:10px 14px; cursor:pointer; margin:0; display:inline-flex; align-items:center; justify-content:center; min-height:40px; line-height:1.2; }
    button.alt { background:#444; }
    button.danger { background:#b03a2e; }
    .muted { color:#aaa; }
    .mono { font-family: Consolas, monospace; }
    .status { margin-top:10px; font-weight:600; }
    .status-grid { display:grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap:12px; }
    .status-box { background:#101010; border:1px solid #383838; border-radius:10px; padding:12px; }
    .status-row { display:flex; align-items:center; justify-content:space-between; gap:10px; }
    .status-dot { width:10px; height:10px; border-radius:999px; display:inline-block; margin-right:8px; }
    .online { background:#1f9d55; }
    .offline { background:#c0392b; }
    .notset { background:#666; }
    .status-actions { margin-top:10px; display:flex; gap:8px; flex-wrap:wrap; }
    .action-row { display:grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap:10px; align-items:stretch; }
    .action-row button { width:100%; }
    .tiny { padding:6px 10px; font-size:12px; }
    .sat-row { border-top:1px solid #303030; padding-top:10px; margin-top:10px; }
    .firmware-current { border-left:4px solid #1f9d55; }
    .firmware-outdated { border-left:4px solid #c0392b; }
    .firmware-ahead { border-left:4px solid #d99a1f; }
    .firmware-unknown { border-left:4px solid #666; }
    .badge { display:inline-block; border-radius:999px; padding:2px 8px; font-size:12px; margin-left:8px; }
    .badge-current { background:#1f9d55; color:#fff; }
    .badge-outdated { background:#c0392b; color:#fff; }
    .badge-ahead { background:#d99a1f; color:#111; }
    .badge-unknown { background:#666; color:#fff; }
  </style>
</head>
<body>
  <div class="wrap">
    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px;">
      <div>
        <h1 style="margin:0;">HubVoiceSat Setup</h1>
        <p id="build_meta" class="muted mono" style="margin:4px 0 0;">Firmware: ...</p>
      </div>
      <button onclick="openControlDeck()" style="background:#1f9d55;font-size:1rem;padding:10px 18px;">HubMusic</button>
    </div>

    <div class="card">
      <h2>Wi-Fi</h2>
      <div class="grid">
        <div>
          <label for="wifi_ssid">Wi-Fi SSID</label>
          <input id="wifi_ssid" placeholder="MyWiFi">
        </div>
        <div>
          <label for="wifi_password">Wi-Fi Password</label>
          <input id="wifi_password" type="password" placeholder="password">
        </div>
      </div>
    </div>

    <div class="card">
      <h2>HubVoice</h2>
      <div class="grid">
        <div>
          <label for="hubvoice_url">HubVoice URL</label>
          <input id="hubvoice_url" placeholder="http://192.168.4.23:8080">
        </div>
        <div>
          <label for="callback_url">Callback URL</label>
          <input id="callback_url" placeholder="http://192.168.4.23:8080/answer">
        </div>
        <div>
          <label for="hubitat_host">Hubitat Host URL</label>
          <input id="hubitat_host" placeholder="http://192.168.4.141">
        </div>
        <div>
          <label for="hubitat_app_id">Hubitat App ID</label>
          <input id="hubitat_app_id" placeholder="8745">
        </div>
        <div class="full">
          <label for="hubitat_access_token">Hubitat Access Token</label>
          <input id="hubitat_access_token" type="password" placeholder="access token">
        </div>
        <div class="full">
          <label for="piper_voice_model">Piper Voice Model</label>
          <select id="piper_voice_model"></select>
          <p class="muted" style="margin-top:6px;">Choose another installed voice and click Save to switch to it.</p>
        </div>
      </div>
    </div>

    <div class="card">
      <h2>Satellites</h2>
      <label for="satellites_text">Known satellites</label>
      <textarea id="satellites_text" placeholder="sat-lr,192.168.4.135,Living Room"></textarea>
      <p class="muted">One line per satellite in the format <span class="mono">id,ip[,alias]</span>. Alias is optional.</p>
    </div>

    <div class="card">
      <div class="action-row">
        <button onclick="saveSetup()">Save</button>
        <button class="alt" onclick="loadSetup()">Reload</button>
        <button class="alt" onclick="refreshStatus()">Refresh Status</button>
        <button class="alt" onclick="createDesktopShortcut()">Create Desktop Shortcut</button>
        <button class="danger" onclick="shutdownEverything()">Shut Down Everything</button>
      </div>
      <div id="status" class="status"></div>
    </div>

    <div class="card">
      <div class="status-row">
        <h2>Status & Controls</h2>
        <label class="muted" style="display:flex;align-items:center;gap:8px;">
          <input id="auto_refresh" type="checkbox" checked style="width:auto;">
          Auto refresh
        </label>
      </div>
      <div id="status_grid" class="status-grid"></div>
      <div id="satellite_status" style="margin-top:12px;"></div>
    </div>
  </div>

  <script>
    let refreshTimer = null;
    let cachedState = null;

    function openControlDeck() {
      const url = (cachedState && cachedState.hubvoice_url) || 'http://127.0.0.1:8080';
      window.open(url.replace(/\/$/, '') + '/control', '_blank');
    }

    function statusClassFromValue(value) {
      if (value === 'online') return 'online';
      if (value === 'offline') return 'offline';
      return 'notset';
    }

    function statusLabelFromValue(value) {
      if (value === 'online') return 'Online';
      if (value === 'offline') return 'Offline';
      return 'Not set';
    }

    function escapeHtml(value) {
      return String(value || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
    }

    function renderStatusCard(title, item, openUrl, extraButtons) {
      const statusValue = item && item.status ? item.status : 'not_set';
      const statusCss = statusClassFromValue(statusValue);
      const statusLabel = statusLabelFromValue(statusValue);
      const detail = item && item.detail ? item.detail : '';
      const openButton = openUrl ? `<button class="alt tiny" onclick="openUrl('${openUrl}')">Open</button>` : '';
      const actions = [openButton, extraButtons || ''].filter(Boolean).join('');
      return `
        <div class="status-box">
          <div class="status-row">
            <div><span class="status-dot ${statusCss}"></span>${escapeHtml(title)}</div>
            <div>${statusLabel}</div>
          </div>
          <p class="muted" style="margin:10px 0 0;">${escapeHtml(detail)}</p>
          <div class="status-actions">${actions}</div>
        </div>
      `;
    }

    async function openUrl(url) {
      const res = await fetch('/api/open_url', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ url })
      });
      const data = await res.json();
      document.getElementById('status').textContent = data.message || (data.ok ? 'Opened.' : 'Failed.');
    }

    function scheduleStatusRefresh() {
      if (refreshTimer) {
        clearInterval(refreshTimer);
        refreshTimer = null;
      }
      if (document.getElementById('auto_refresh').checked) {
        refreshTimer = setInterval(refreshStatus, 10000);
      }
    }

    async function refreshStatus() {
      const res = await fetch('/api/status');
      const data = await res.json();
      const urls = data.urls || {};

      const gridHtml = [
        renderStatusCard('Setup Page', data.setup_page, urls.setup_page || 'http://127.0.0.1:8093/'),
        renderStatusCard('HubVoice URL', data.hubvoice, urls.hubvoice || ''),
        renderStatusCard('Voice Pipeline', data.hubvoice_runtime, urls.hubvoice || ''),
        renderStatusCard('HubVoice Port', data.hubvoice_port, urls.hubvoice || ''),
        renderStatusCard('Callback URL', data.callback, urls.callback || ''),
        renderStatusCard('Callback Port', data.callback_port, urls.callback || ''),
        renderStatusCard('Hubitat Health', data.hubitat_health, '')
      ];
      document.getElementById('status_grid').innerHTML = gridHtml.join('');

      const satellites = Array.isArray(data.satellites) ? data.satellites : [];
      if (!satellites.length) {
        document.getElementById('satellite_status').innerHTML = '<p class="muted">No satellites configured.</p>';
      } else {
        const satHtml = satellites.map(sat => {
          const state = sat.update_status || 'unknown';
          const rowClass = state === 'current' ? 'firmware-current' : state === 'outdated' ? 'firmware-outdated' : state === 'ahead' ? 'firmware-ahead' : 'firmware-unknown';
          const badgeClass = state === 'current' ? 'badge-current' : state === 'outdated' ? 'badge-outdated' : state === 'ahead' ? 'badge-ahead' : 'badge-unknown';
          const badgeText = state === 'current' ? 'Current' : state === 'outdated' ? 'Update Needed' : state === 'ahead' ? 'Ahead' : 'Unknown';

          return `
          <div class="status-box sat-row ${rowClass}">
            <div class="status-row">
              <div><strong>${escapeHtml(sat.alias || sat.name)}</strong> <span class="muted">(ID: ${escapeHtml(sat.name)} | ${escapeHtml(sat.ip)})</span><span class="badge ${badgeClass}">${badgeText}</span></div>
              <button class="alt tiny" onclick="openUrl('http://${escapeHtml(sat.ip)}${sat.web_port ? ':' + sat.web_port : ''}/')">Open ${sat.web_port ? ':' + sat.web_port : ''}</button>
            </div>
            <p class="muted" style="margin:10px 0 0;">
              Ping: ${sat.ping ? 'Online' : 'Offline'} |
              Web: ${sat.web ? 'Online' : 'Offline'}${sat.web_port ? ' on port ' + sat.web_port : ''}
            </p>
            <p class="mono muted" style="margin:8px 0 0;">
              URL: http://${escapeHtml(sat.ip)}${sat.web_port ? ':' + sat.web_port : ''}/
            </p>
          </div>
        `;
        }).join('');
        document.getElementById('satellite_status').innerHTML = satHtml;
      }
    }

    function loadVoiceOptions(selectedValue, options) {
      const select = document.getElementById('piper_voice_model');
      select.innerHTML = '';

      const values = Array.isArray(options) ? options.slice() : [];
      if (selectedValue && !values.includes(selectedValue)) {
        values.unshift(selectedValue);
      }

      for (const value of values) {
        const option = document.createElement('option');
        option.value = value;
        option.textContent = value;
        if (value === selectedValue) option.selected = true;
        select.appendChild(option);
      }
    }

    async function loadSetup() {
      const res = await fetch('/api/state');
      const data = await res.json();
      cachedState = data;

      const launcherVersion = data.launcher_version || 'unknown';
      const firmwareVersion = data.firmware_target_version || 'unknown';
      document.getElementById('build_meta').textContent = `Firmware: ${firmwareVersion}`;

      for (const [key, value] of Object.entries(data)) {
        const el = document.getElementById(key);
        if (el) el.value = value || '';
      }
      loadVoiceOptions(data.piper_voice_model || '', data.piper_voice_models || []);
      const tokenInput = document.getElementById('hubitat_access_token');
      tokenInput.value = '';
      tokenInput.placeholder = data.hubitat_access_token_saved ? '********' : 'access token';
      document.getElementById('status').textContent = '';
      await refreshStatus();
    }

    async function saveSetup() {
      const payload = {
        wifi_ssid: document.getElementById('wifi_ssid').value.trim(),
        wifi_password: document.getElementById('wifi_password').value,
        hubvoice_url: document.getElementById('hubvoice_url').value.trim(),
        hubitat_host: document.getElementById('hubitat_host').value.trim(),
        hubitat_app_id: document.getElementById('hubitat_app_id').value.trim(),
        hubitat_access_token: document.getElementById('hubitat_access_token').value.trim(),
        callback_url: document.getElementById('callback_url').value.trim(),
        piper_voice_model: document.getElementById('piper_voice_model').value.trim(),
        satellites_text: document.getElementById('satellites_text').value
      };

      const res = await fetch('/api/save', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const data = await res.json();
      document.getElementById('status').textContent = data.message || (data.ok ? 'Saved.' : 'Failed.');
      cachedState = null;
      await loadSetup();
    }

    async function createDesktopShortcut() {
      const res = await fetch('/api/create_desktop_shortcut', {
        method: 'POST'
      });
      const data = await res.json();
      document.getElementById('status').textContent = data.message || (data.ok ? 'Shortcut created.' : 'Failed.');
    }

    async function shutdownEverything() {
      const ok = confirm('Shut down setup server, runtime, and launcher now?');
      if (!ok) return;

      const res = await fetch('/api/shutdown_all', {
        method: 'POST'
      });
      const data = await res.json();
      document.getElementById('status').textContent = data.message || (data.ok ? 'Shutting down everything.' : 'Failed.');
    }

    document.getElementById('auto_refresh').addEventListener('change', scheduleStatusRefresh);
    loadSetup();
    scheduleStatusRefresh();
  </script>
</body>
</html>
'@

function Start-RuntimeIfNeeded {
  $runtimeVenvPythonw = Join-Path $root ".runtime-venv\Scripts\pythonw.exe"
  $runtimeVenvPython = Join-Path $root ".runtime-venv\Scripts\python.exe"
  $venvPythonw = Join-Path $root ".venv\Scripts\pythonw.exe"
  $venvPython = Join-Path $root ".venv\Scripts\python.exe"
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
