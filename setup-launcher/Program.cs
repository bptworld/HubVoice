using System.Diagnostics;
using System.Drawing;
using System.Net.Http;
using System.Net.Sockets;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

internal static class Program
{
    private const int PreferredSetupPort = 8093;
    private const string RequiredSetupSchemaVersion = "2";
    private const string PayloadPrefix = "Payload/";
    private const string LauncherPathEnvVar = "HUBVOICESAT_LAUNCHER_PATH";
    private static readonly Mutex SingleInstanceMutex = new(false, @"Local\HubVoiceSatSetupLauncher");
    private static readonly HttpClient HttpClientInstance = new() 
    { 
        Timeout = TimeSpan.FromSeconds(2) 
    };

    [STAThread]
    private static async Task<int> Main()
    {
        var currentExePath = Environment.ProcessPath;
        if (!string.IsNullOrWhiteSpace(currentExePath) && TryForwardToNewerLauncher(currentExePath))
        {
            return 0;
        }

        var lockTaken = false;
        try
        {
            try
            {
                lockTaken = SingleInstanceMutex.WaitOne(0, false);
            }
            catch (AbandonedMutexException)
            {
                // If the previous launcher crashed, recover ownership and continue.
                lockTaken = true;
            }

            if (!lockTaken)
            {
                var existingSetupUrl = await FindCompatibleSetupUrlAsync() ?? BuildSetupUrl(PreferredSetupPort);
                OpenBrowser(existingSetupUrl);
                return 0;
            }

            await EnsureLocalRuntimeReadyAsync();
            var scriptPath = EnsurePayloadReady();

            var setupUrl = await FindCompatibleSetupUrlAsync();
            if (setupUrl is null)
            {
                StopExistingSetupServer(scriptPath);
                var setupPort = await ResolveSetupPortAsync();
                setupUrl = BuildSetupUrl(setupPort);
                StartSetupServer(scriptPath, setupPort);

                if (!await WaitForSetupAsync(setupUrl, TimeSpan.FromSeconds(20)))
                {
                    ShowError("Failed to start the HubVoiceSat setup page.");
                    return 1;
                }
            }

            OpenBrowser(setupUrl);

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            using var trayContext = new SetupTrayContext(setupUrl);
            Application.Run(trayContext);
            return 0;
        }
        catch (Exception ex)
        {
            ShowError($"HubVoiceSat setup failed.{Environment.NewLine}{Environment.NewLine}{ex.Message}");
            return 1;
        }
        finally
        {
            if (lockTaken)
            {
                try
                {
                    SingleInstanceMutex.ReleaseMutex();
                }
                catch
                {
                }
            }

            HttpClientInstance?.Dispose();
            SingleInstanceMutex?.Dispose();
        }
    }

    private static bool TryForwardToNewerLauncher(string currentExePath)
    {
        try
        {
            var currentVersion = GetLauncherVersion(currentExePath);
            if (currentVersion is null)
            {
                return false;
            }

            var newerPath = FindNewestLauncherPath(currentExePath, currentVersion);
            if (string.IsNullOrWhiteSpace(newerPath))
            {
                return false;
            }

            var psi = new ProcessStartInfo
            {
                FileName = newerPath,
                WorkingDirectory = Path.GetDirectoryName(newerPath) ?? AppContext.BaseDirectory,
                UseShellExecute = true,
            };
            Process.Start(psi);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static string? FindNewestLauncherPath(string currentExePath, Version currentVersion)
    {
        var candidates = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var currentDir = Path.GetDirectoryName(currentExePath) ?? AppContext.BaseDirectory;
        var currentDirInfo = new DirectoryInfo(currentDir);

        AddLauncherCandidate(candidates, Path.Combine(currentDir, "HubVoiceSat.exe"));
        AddLauncherCandidate(candidates, Path.Combine(currentDir, "HubVoiceSatSetup.exe"));

        // If launched from repo root, scan releases for newer launchers.
        var releasesUnderCurrent = Path.Combine(currentDir, "releases");
        AddLaunchersFromReleases(candidates, releasesUnderCurrent);

        // If launched from a release folder, scan sibling releases + repo root launchers.
        var parentDirInfo = currentDirInfo.Parent;
        if (parentDirInfo is not null && parentDirInfo.Name.Equals("releases", StringComparison.OrdinalIgnoreCase))
        {
            AddLaunchersFromReleases(candidates, parentDirInfo.FullName);
            var repoRoot = parentDirInfo.Parent?.FullName;
            if (!string.IsNullOrWhiteSpace(repoRoot))
            {
                AddLauncherCandidate(candidates, Path.Combine(repoRoot, "HubVoiceSat.exe"));
                AddLauncherCandidate(candidates, Path.Combine(repoRoot, "HubVoiceSatSetup.exe"));
            }
        }

        Version bestVersion = currentVersion;
        string? bestPath = null;
        foreach (var candidate in candidates)
        {
            if (candidate.Equals(currentExePath, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var candidateVersion = GetLauncherVersion(candidate);
            if (candidateVersion is null)
            {
                continue;
            }

            if (candidateVersion > bestVersion)
            {
                bestVersion = candidateVersion;
                bestPath = candidate;
            }
        }

        return bestPath;
    }

    private static void AddLaunchersFromReleases(HashSet<string> candidates, string releasesDir)
    {
        if (!Directory.Exists(releasesDir))
        {
            return;
        }

        foreach (var dir in Directory.EnumerateDirectories(releasesDir, "hubvoice-sat-*-release"))
        {
            AddLauncherCandidate(candidates, Path.Combine(dir, "HubVoiceSat.exe"));
            AddLauncherCandidate(candidates, Path.Combine(dir, "HubVoiceSatSetup.exe"));
        }
    }

    private static void AddLauncherCandidate(HashSet<string> candidates, string path)
    {
        if (File.Exists(path))
        {
            candidates.Add(path);
        }
    }

    private static Version? GetLauncherVersion(string path)
    {
        if (!File.Exists(path))
        {
            return null;
        }

        var fileVersion = FileVersionInfo.GetVersionInfo(path).FileVersion;
        if (string.IsNullOrWhiteSpace(fileVersion))
        {
            return null;
        }

        var match = Regex.Match(fileVersion, "(\\d+)\\.(\\d+)\\.(\\d+)\\.(\\d+)");
        if (!match.Success)
        {
            return null;
        }

        if (!int.TryParse(match.Groups[1].Value, out var major) ||
            !int.TryParse(match.Groups[2].Value, out var minor) ||
            !int.TryParse(match.Groups[3].Value, out var build) ||
            !int.TryParse(match.Groups[4].Value, out var revision))
        {
            return null;
        }

        return new Version(major, minor, build, revision);
    }

    private static async Task<int> ResolveSetupPortAsync()
    {
        for (var port = PreferredSetupPort; port <= PreferredSetupPort + 20; port++)
        {
            if (!IsPortOpen(port))
            {
                return port;
            }
        }

        return PreferredSetupPort;
    }

    private static async Task EnsureLocalRuntimeReadyAsync()
    {
        var root = AppContext.BaseDirectory;
        var runtimeExe = Path.Combine(root, "HubVoiceRuntime.exe");
        var runtimeScript = Path.Combine(root, "hubvoice-runtime.py");
        var runtimePort = GetRuntimePort(root);
        if (!File.Exists(runtimeExe) && !File.Exists(runtimeScript))
        {
            return;
        }

        if (!File.Exists(runtimeExe))
        {
            EnsureRuntimeVenv(root);
        }

        StopExistingLocalRuntime(root, runtimePort);
        StartLocalRuntime(root, runtimeScript, runtimeExe);

        if (!await WaitForRuntimeReadyAsync(runtimePort, TimeSpan.FromSeconds(60)))
        {
            throw new InvalidOperationException($"The HubVoiceSat runtime did not start on port {runtimePort}.");
        }
    }

    /// <summary>
    /// If no runtime venv exists yet, runs setup-runtime.ps1 to create it.
    /// Shows a visible PowerShell window so the user can see installation progress.
    /// </summary>
    private static void EnsureRuntimeVenv(string root)
    {
        var venvPython = Path.Combine(root, ".runtime-venv", "Scripts", "python.exe");
        if (File.Exists(venvPython))
        {
            return;
        }

        var setupScript = Path.Combine(root, "setup-runtime.ps1");
        if (!File.Exists(setupScript))
        {
            return;
        }

        _ = System.Windows.Forms.MessageBox.Show(
            "The HubVoice Python runtime needs to be set up before the first launch.\n\n" +
            "A setup window will open and install the required packages.\n" +
            "This takes a few minutes — please wait for it to finish before continuing.",
            "HubVoiceSat — First-Time Setup",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Information);

        var psi = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = $"-NoProfile -ExecutionPolicy Bypass -File \"{setupScript}\"",
            WorkingDirectory = root,
            UseShellExecute = false,
            CreateNoWindow = false,
        };

        using var process = Process.Start(psi)
            ?? throw new InvalidOperationException("Failed to start setup-runtime.ps1.");

        process.WaitForExit();

        if (!File.Exists(venvPython))
        {
            throw new InvalidOperationException(
                "Python runtime setup did not complete successfully.\n" +
                "Make sure Python 3.10 or newer is installed and on PATH, then run setup-runtime.ps1 manually.");
        }
    }

    private static void StopExistingLocalRuntime(string root, int runtimePort)
    {
        var escapedRoot = root.Replace("'", "''");
        var tempScript = Path.ChangeExtension(Path.GetTempFileName(), ".ps1");
        File.WriteAllText(tempScript,
            "$root = '" + escapedRoot + "'\r\n" +
            "Get-CimInstance Win32_Process | " +
            "Where-Object { " +
            "(($_.CommandLine -like ('*hubvoice-runtime.py*')) -and ($_.CommandLine -like ('*' + $root + '*'))) -or " +
            "(($_.Name -eq 'HubVoiceRuntime.exe') -and ($_.ExecutablePath -like ('*' + $root + '*'))) " +
            "} | " +
            "ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }\r\n");

        var psi = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = $"-WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File \"{tempScript}\"",
            UseShellExecute = true,
            WindowStyle = ProcessWindowStyle.Hidden,
        };

        using var process = Process.Start(psi);
        process?.WaitForExit(5000);
        try { File.Delete(tempScript); } catch { }
        WaitForLocalRuntimeToExit(root, runtimePort, TimeSpan.FromSeconds(10));
    }

    private static void StartLocalRuntime(string root, string runtimeScript, string runtimeExe)
    {
        var preferWorkspaceRuntimeScript = FindWorkspaceScript() is not null && File.Exists(runtimeScript);

        if (!preferWorkspaceRuntimeScript && File.Exists(runtimeExe))
        {
            var exePsi = new ProcessStartInfo
            {
                FileName = runtimeExe,
                WorkingDirectory = root,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            Process.Start(exePsi);
            return;
        }

        var runtimeVenvPythonw = Path.Combine(root, ".runtime-venv", "Scripts", "pythonw.exe");
        var runtimeVenvPython = Path.Combine(root, ".runtime-venv", "Scripts", "python.exe");
        var venvPythonw = Path.Combine(root, ".venv", "Scripts", "pythonw.exe");
        var venvPython = Path.Combine(root, ".venv", "Scripts", "python.exe");
        ProcessStartInfo psi;

        if (File.Exists(runtimeVenvPythonw))
        {
            psi = new ProcessStartInfo
            {
                FileName = runtimeVenvPythonw,
                Arguments = $"-u \"{runtimeScript}\"",
                WorkingDirectory = root,
                UseShellExecute = false,
                CreateNoWindow = true
            };
        }
        else if (File.Exists(runtimeVenvPython))
        {
            psi = new ProcessStartInfo
            {
                FileName = runtimeVenvPython,
                Arguments = $"-u \"{runtimeScript}\"",
                WorkingDirectory = root,
                UseShellExecute = false,
                CreateNoWindow = true
            };
        }
        else if (File.Exists(venvPythonw))
        {
            psi = new ProcessStartInfo
            {
                FileName = venvPythonw,
                Arguments = $"-u \"{runtimeScript}\"",
                WorkingDirectory = root,
                UseShellExecute = false,
                CreateNoWindow = true
            };
        }
        else if (File.Exists(venvPython))
        {
            psi = new ProcessStartInfo
            {
                FileName = venvPython,
                Arguments = $"-u \"{runtimeScript}\"",
                WorkingDirectory = root,
                UseShellExecute = false,
                CreateNoWindow = true
            };
        }
        else
        {
            psi = new ProcessStartInfo
            {
                FileName = "py",
                Arguments = $"-3 -u \"{runtimeScript}\"",
                WorkingDirectory = root,
                UseShellExecute = false,
                CreateNoWindow = true
            };
        }

        Process.Start(psi);
    }

    private static void WaitForLocalRuntimeToExit(string root, int runtimePort, TimeSpan timeout)
    {
        var deadline = DateTimeOffset.UtcNow.Add(timeout);
        while (DateTimeOffset.UtcNow < deadline)
        {
            if (!HasMatchingLocalRuntimeProcesses(root))
            {
                break;
            }

            Thread.Sleep(200);
        }

        WaitForPortToClose(runtimePort, timeout);
    }

    private static bool HasMatchingLocalRuntimeProcesses(string root)
    {
        try
        {
            using var process = Process.Start(new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"$root = '{root.Replace("'", "''")}'; $match = Get-CimInstance Win32_Process | Where-Object {{ (($_.CommandLine -like ('*hubvoice-runtime.py*')) -and ($_.CommandLine -like ('*' + $root + '*'))) -or (($_.Name -eq 'HubVoiceRuntime.exe') -and ($_.ExecutablePath -like ('*' + $root + '*'))) }} | Select-Object -First 1; if ($match) {{ 'true' }} else {{ 'false' }}\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true
            });
            var output = process?.StandardOutput.ReadToEnd().Trim();
            process?.WaitForExit(3000);
            return string.Equals(output, "true", StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private static void WaitForPortToClose(int port, TimeSpan timeout)
    {
        var deadline = DateTimeOffset.UtcNow.Add(timeout);
        while (DateTimeOffset.UtcNow < deadline)
        {
            if (!IsPortOpen(port))
            {
                return;
            }

            Thread.Sleep(200);
        }
    }

    private static bool IsPortOpen(int port)
    {
        try
        {
            return System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties()
                .GetActiveTcpListeners()
                .Any(endpoint => endpoint.Port == port);
        }
        catch
        {
            return false;
        }
    }

    private static int GetRuntimePort(string root)
    {
        var configPath = Path.Combine(root, "hubvoice-sat-setup.json");
        if (!File.Exists(configPath))
        {
            return 8080;
        }

        try
        {
            using var doc = JsonDocument.Parse(File.ReadAllText(configPath));
            foreach (var propertyName in new[] { "hubvoice_url", "callback_url" })
            {
                if (!doc.RootElement.TryGetProperty(propertyName, out var urlElement))
                {
                    continue;
                }

                var rawUrl = urlElement.GetString();
                if (!string.IsNullOrWhiteSpace(rawUrl) &&
                    Uri.TryCreate(rawUrl, UriKind.Absolute, out var uri) &&
                    !uri.IsDefaultPort)
                {
                    return uri.Port;
                }
            }
        }
        catch
        {
        }

        return 8080;
    }

    private static string EnsurePayloadReady()
    {
        // If the EXE is running from within the workspace directory tree, prefer the
        // live workspace copy of setup-web.ps1 so changes take effect without a rebuild.
        var workspaceScript = FindWorkspaceScript();
        if (workspaceScript is not null)
        {
            return workspaceScript;
        }

        var payloadRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "HubVoiceSatSetup");

        Directory.CreateDirectory(payloadRoot);
        ExtractPayload(payloadRoot);

        var scriptPath = Path.Combine(payloadRoot, "setup-web.ps1");
        if (!File.Exists(scriptPath))
        {
            throw new FileNotFoundException("The embedded setup-web.ps1 payload was not extracted.", scriptPath);
        }

        return scriptPath;
    }

    /// <summary>
    /// Walks up from the EXE's directory looking for setup-web.ps1.
    /// Returns the path if found (meaning we're running from the workspace tree),
    /// or null if the EXE is running standalone (e.g. from a release package).
    /// </summary>
    private static string? FindWorkspaceScript()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        for (var depth = 0; depth < 6 && dir is not null; depth++, dir = dir.Parent)
        {
            var candidate = Path.Combine(dir.FullName, "setup-web.ps1");
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        return null;
    }

    private static string BuildSetupUrl(int port)
    {
        return $"http://127.0.0.1:{port}/";
    }

    private static async Task<string?> FindCompatibleSetupUrlAsync()
    {
        for (var port = PreferredSetupPort; port <= PreferredSetupPort + 20; port++)
        {
            if (!IsPortOpen(port))
            {
                continue;
            }

            var candidateUrl = BuildSetupUrl(port);
            if (await IsSetupApiCompatibleAsync(candidateUrl))
            {
                return candidateUrl;
            }
        }

        return null;
    }

    private static void StartSetupServer(string scriptPath, int setupPort)
    {
        var launcherPath = (Environment.ProcessPath ?? Application.ExecutablePath).Replace("'", "''");
        var escapedScript = scriptPath.Replace("'", "''");

        // Write a tiny wrapper script so we can pass env vars without -Command quoting issues.
        var tempWrapper = Path.ChangeExtension(Path.GetTempFileName(), ".ps1");
        File.WriteAllText(tempWrapper,
            $"$env:{LauncherPathEnvVar} = '{launcherPath}'\r\n" +
            $"$env:HUBVOICESAT_SETUP_PORT = '{setupPort}'\r\n" +
            $"& '{escapedScript}'\r\n");

        var psi = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = $"-WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File \"{tempWrapper}\"",
            UseShellExecute = true,
            WindowStyle = ProcessWindowStyle.Hidden,
            WorkingDirectory = Path.GetDirectoryName(scriptPath) ?? AppContext.BaseDirectory
        };

        Process.Start(psi);
    }

    private static async Task<bool> WaitForPortAsync(string host, int port, TimeSpan timeout)
    {
        var deadline = DateTimeOffset.UtcNow.Add(timeout);
        while (DateTimeOffset.UtcNow < deadline)
        {
            if (await IsPortOpenAsync(host, port))
            {
                return true;
            }

            await Task.Delay(500);
        }

        return false;
    }

    private static async Task<bool> IsPortOpenAsync(string host, int port)
    {
        try
        {
            using var client = new TcpClient();
            var connectTask = client.ConnectAsync(host, port);
            var completed = await Task.WhenAny(connectTask, Task.Delay(1000));
            return completed == connectTask && client.Connected;
        }
        catch
        {
            return false;
        }
    }

    private static async Task<bool> WaitForRuntimeReadyAsync(int port, TimeSpan timeout)
    {
        var deadline = DateTimeOffset.UtcNow.Add(timeout);
        while (DateTimeOffset.UtcNow < deadline)
        {
            if (await IsRuntimeHealthyAsync(port))
            {
                return true;
            }

            await Task.Delay(500);
        }

        return false;
    }

    private static async Task<bool> IsRuntimeHealthyAsync(int port)
    {
        try
        {
            using var response = await HttpClientInstance.GetAsync($"http://127.0.0.1:{port}/");
            if (!response.IsSuccessStatusCode)
            {
                return false;
            }

            using var stream = await response.Content.ReadAsStreamAsync();
            using var doc = await JsonDocument.ParseAsync(stream);
            if (!doc.RootElement.TryGetProperty("ok", out var okElement) || !okElement.GetBoolean())
            {
                return false;
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    private static void StopExistingSetupServer(string scriptPath)
    {
        // Use a temp script file to avoid any quoting issues with -Command on the command line.
        var tempScript = Path.ChangeExtension(Path.GetTempFileName(), ".ps1");
        File.WriteAllText(tempScript,
            "Get-CimInstance Win32_Process -Filter \"Name = 'powershell.exe'\" | " +
            "Where-Object { $_.CommandLine -like '*setup-web.ps1*' } | " +
            "ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }");

        var psi = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = $"-WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File \"{tempScript}\"",
            UseShellExecute = true,
            WindowStyle = ProcessWindowStyle.Hidden,
        };

        using var process = Process.Start(psi);
        process?.WaitForExit(5000);
        Thread.Sleep(500);

        try { File.Delete(tempScript); } catch { }
    }

    private static void ExtractPayload(string payloadRoot)
    {
        var assembly = Assembly.GetExecutingAssembly();
        var resourceNames = assembly.GetManifestResourceNames()
            .Where(name => name.StartsWith(PayloadPrefix, StringComparison.Ordinal))
            .ToArray();

        if (resourceNames.Length == 0)
        {
            throw new InvalidOperationException("The setup payload is missing from the executable.");
        }

        foreach (var resourceName in resourceNames)
        {
            var relativePath = resourceName[PayloadPrefix.Length..].Replace('/', Path.DirectorySeparatorChar);
            var destinationPath = Path.Combine(payloadRoot, relativePath);
            var destinationDir = Path.GetDirectoryName(destinationPath);
            if (!string.IsNullOrEmpty(destinationDir))
            {
                Directory.CreateDirectory(destinationDir);
            }

            if (IsUserManagedFile(relativePath) && File.Exists(destinationPath))
            {
                continue;
            }

            using var sourceStream = assembly.GetManifestResourceStream(resourceName)
                ?? throw new InvalidOperationException($"Missing embedded payload resource '{resourceName}'.");
            using var destinationStream = File.Create(destinationPath);
            sourceStream.CopyTo(destinationStream);
        }
    }

    private static bool IsUserManagedFile(string relativePath)
    {
        return relativePath.Equals("hubvoice-sat-setup.json", StringComparison.OrdinalIgnoreCase) ||
               relativePath.Equals("secrets.yaml", StringComparison.OrdinalIgnoreCase) ||
               relativePath.Equals("satellites.csv", StringComparison.OrdinalIgnoreCase);
    }

    private static async Task<bool> WaitForSetupAsync(string setupUrl, TimeSpan timeout)
    {
        var deadline = DateTimeOffset.UtcNow.Add(timeout);
        while (DateTimeOffset.UtcNow < deadline)
        {
            if (await IsSetupReadyAsync(setupUrl))
            {
                return true;
            }

            await Task.Delay(500);
        }

        return false;
    }

    private static async Task<bool> IsSetupReadyAsync(string setupUrl)
    {
        try
        {
            using var response = await HttpClientInstance.GetAsync(setupUrl);
            var code = (int)response.StatusCode;
            if (code < 200 || code >= 500)
            {
                return false;
            }

            return await IsSetupApiCompatibleAsync(setupUrl);
        }
        catch
        {
            return false;
        }
    }

    private static async Task<bool> IsSetupApiCompatibleAsync(string setupUrl)
    {
        try
        {
            using var response = await HttpClientInstance.GetAsync(setupUrl + "api/state");
            if (!response.IsSuccessStatusCode)
            {
                return false;
            }

            using var stream = await response.Content.ReadAsStreamAsync();
            using var doc = await JsonDocument.ParseAsync(stream);
            if (!doc.RootElement.TryGetProperty("launcher_version", out _)
                || !doc.RootElement.TryGetProperty("firmware_target_version", out _)
                || !doc.RootElement.TryGetProperty("setup_schema_version", out var schemaElement))
            {
                return false;
            }

            var schemaVersion = schemaElement.ValueKind == JsonValueKind.String
                ? schemaElement.GetString()
                : schemaElement.ToString();

            return string.Equals(schemaVersion, RequiredSetupSchemaVersion, StringComparison.Ordinal);
        }
        catch
        {
            return false;
        }
    }

    private static async Task<bool> RequestSetupShutdownAsync(string setupUrl)
    {
        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Post, setupUrl + "api/shutdown");
            using var response = await HttpClientInstance.SendAsync(request);
            return response.IsSuccessStatusCode;
        }
        catch
        {
            return false;
        }
    }

    private sealed class SetupTrayContext : ApplicationContext
    {
        private readonly string _setupUrl;
        private readonly NotifyIcon _notifyIcon;
        private readonly ToolStripMenuItem _statusItem;
        private readonly System.Windows.Forms.Timer _statusTimer;
        private readonly Icon _trayIcon;

        public SetupTrayContext(string setupUrl)
        {
            _setupUrl = setupUrl;
            _statusItem = new ToolStripMenuItem("Status: checking") { Enabled = false };

            var menu = new ContextMenuStrip();
            menu.Items.Add(_statusItem);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(new ToolStripMenuItem("Open Setup", null, (_, _) => OpenBrowser(_setupUrl)));
            menu.Items.Add(new ToolStripMenuItem("Shut Down Server", null, async (_, _) => await OnShutdownClickedAsync()));
            menu.Items.Add(new ToolStripMenuItem("Shutdown + Exit", null, async (_, _) => await OnShutdownAndExitClickedAsync()));
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(new ToolStripMenuItem("Exit Tray", null, (_, _) => ExitThread()));

            _trayIcon = GetPreferredTrayIcon();

            _notifyIcon = new NotifyIcon
            {
                Icon = _trayIcon,
                Text = "HubVoiceSat Setup",
                Visible = true,
                ContextMenuStrip = menu
            };
            _notifyIcon.DoubleClick += (_, _) => OpenBrowser(_setupUrl);

            _statusTimer = new System.Windows.Forms.Timer { Interval = 4000 };
            _statusTimer.Tick += async (_, _) => await UpdateStatusAsync();
            _statusTimer.Start();

            _ = UpdateStatusAsync();
        }

        private async Task OnShutdownClickedAsync()
        {
            await RequestSetupShutdownAsync(_setupUrl);
            await Task.Delay(400);
            await UpdateStatusAsync();
        }

        private async Task OnShutdownAndExitClickedAsync()
        {
            await RequestSetupShutdownAsync(_setupUrl);
            ExitThread();
        }

        private async Task UpdateStatusAsync()
        {
            var running = await IsSetupApiCompatibleAsync(_setupUrl);
            _statusItem.Text = running ? "Status: running" : "Status: stopped";
            _notifyIcon.Text = running ? "HubVoiceSat Setup (running)" : "HubVoiceSat Setup (stopped)";
        }

        protected override void ExitThreadCore()
        {
            _statusTimer.Stop();
            _notifyIcon.Visible = false;
            _notifyIcon.Dispose();
            _statusTimer.Dispose();
            _trayIcon.Dispose();
            base.ExitThreadCore();
        }

        private static Icon GetPreferredTrayIcon()
        {
            try
            {
                var speechIconPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                    "System32",
                    "Speech",
                    "SpeechUX",
                    "SpeechUXWiz.exe");

                if (File.Exists(speechIconPath))
                {
                    var icon = ExtractIndexedIcon(speechIconPath, 0);
                    if (icon is not null)
                    {
                        return icon;
                    }
                }
            }
            catch
            {
            }

            return (Icon)SystemIcons.Application.Clone();
        }

        private static Icon? ExtractIndexedIcon(string filePath, int iconIndex)
        {
            var largeIcons = new IntPtr[1];
            var smallIcons = new IntPtr[1];
            try
            {
                var extracted = NativeMethods.ExtractIconEx(filePath, iconIndex, largeIcons, smallIcons, 1);
                if (extracted <= 0)
                {
                    return null;
                }

                var iconHandle = largeIcons[0] != IntPtr.Zero ? largeIcons[0] : smallIcons[0];
                if (iconHandle == IntPtr.Zero)
                {
                    return null;
                }

                using var icon = Icon.FromHandle(iconHandle);
                return (Icon)icon.Clone();
            }
            finally
            {
                if (largeIcons[0] != IntPtr.Zero)
                {
                    NativeMethods.DestroyIcon(largeIcons[0]);
                }

                if (smallIcons[0] != IntPtr.Zero)
                {
                    NativeMethods.DestroyIcon(smallIcons[0]);
                }
            }
        }
    }

    private static class NativeMethods
    {
        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        internal static extern uint ExtractIconEx(
            string lpszFile,
            int nIconIndex,
            IntPtr[]? phiconLarge,
            IntPtr[]? phiconSmall,
            uint nIcons);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool DestroyIcon(IntPtr hIcon);
    }

    private static void OpenBrowser(string url)
    {
        var psi = new ProcessStartInfo
        {
            FileName = url,
            UseShellExecute = true
        };

        Process.Start(psi);
    }

    private static void ShowError(string message)
    {
        _ = System.Windows.Forms.MessageBox.Show(
            message,
            "HubVoiceSat Setup",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Error);
    }
}
