using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SnipeAgent
{
    static class AgentInstaller
    {
        public static Task<bool> RunInstallAsync(
            string serverUrl, int pollInterval,
            Action<string> onOutput, CancellationToken ct = default)
            => RunAsync(serverUrl, pollInterval, uninstall: false, onOutput, ct);

        public static Task<bool> RunUninstallAsync(
            Action<string> onOutput, CancellationToken ct = default)
            => RunAsync(string.Empty, 0, uninstall: true, onOutput, ct);

        private static int _lastFlushedLine = 0;

        private static async Task<bool> RunAsync(
            string serverUrl, int pollInterval, bool uninstall,
            Action<string> onOutput, CancellationToken ct)
        {
            // CommonApplicationData = C:\ProgramData — writable by elevated AND normal processes
            string tempDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                "DiskHealthAgentInstall");
            Directory.CreateDirectory(tempDir);

            // Always re-extract so we never run a stale cached version
            AgentScripts.ExtractTo(tempDir);

            string installPath = Path.Combine(tempDir, "InstallDiskAgent.ps1");
            string logPath     = Path.Combine(tempDir, "install_out.txt");
            string exitMarker  = Path.Combine(tempDir, "exit_code.txt");

            File.WriteAllText(logPath,    string.Empty);
            File.WriteAllText(exitMarker, "-1");
            _lastFlushedLine = 0;

            onOutput($"Scripts extracted to: {tempDir}");
            if (!uninstall) onOutput($"Server: {serverUrl}  |  Poll: {pollInterval}s");
            onOutput(uninstall ? "Launching uninstaller..." : "Launching installer — UAC prompt will appear, click Yes...\n");

            string installerArgs = uninstall
                ? "-Uninstall"
                : $"-ServerUrl \"{serverUrl}\" -PollInterval {pollInterval}";

            // The installer writes its own log to logPath via *>&1 > logFile,
            // then writes its exit code to exitMarker.
            // We run it elevated via -Verb runas from a NON-interactive wrapper
            // so UAC fires once, and both wrapper+installer share the same session.
            string wrapPath = Path.Combine(tempDir, "run.ps1");

            // Escape backslashes for use inside a PS double-quoted string
            string logPathEsc  = logPath.Replace("\\", "\\\\");
            string exitPathEsc = exitMarker.Replace("\\", "\\\\");
            string instPathEsc = installPath.Replace("\\", "\\\\");

            string wrapContent =
                "$ErrorActionPreference = 'Continue'\r\n" +
                $"$out = & powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass" +
                $" -File \"{instPathEsc}\" {installerArgs} 2>&1\r\n" +
                $"$code = $LASTEXITCODE\r\n" +
                $"$out | ForEach-Object {{ $_.ToString() }} |" +
                $" Out-File -FilePath \"{logPathEsc}\" -Encoding UTF8 -Append\r\n" +
                $"[System.IO.File]::WriteAllText(\"{exitPathEsc}\", $code)\r\n";

            File.WriteAllText(wrapPath, wrapContent, new UTF8Encoding(true));

            // Elevate the wrapper via runas — it then calls the installer normally
            var psi = new ProcessStartInfo
            {
                FileName         = "powershell.exe",
                Arguments        = $"-NoProfile -ExecutionPolicy Bypass -File \"{wrapPath}\"",
                Verb             = "runas",
                UseShellExecute  = true,   // required for Verb=runas
                WorkingDirectory = tempDir,
            };

            Process? proc;
            try   { proc = Process.Start(psi); }
            catch (Exception ex)
            {
                onOutput("✗ Could not launch elevated installer: " + ex.Message);
                if (ex.Message.Contains("cancel") || ex.Message.Contains("denied"))
                    onOutput("  UAC was cancelled or access was denied.");
                return false;
            }

            if (proc is null) { onOutput("✗ Process did not start."); return false; }

            // Tail the log file while the elevated process runs
            while (!proc.HasExited)
            {
                await Task.Delay(500, ct);
                FlushLog(logPath, onOutput);
            }
            await Task.Delay(300); // let final writes complete
            FlushLog(logPath, onOutput);

            int exitCode = -1;
            try { exitCode = int.Parse(File.ReadAllText(exitMarker).Trim()); } catch { }
            if (exitCode == -1) exitCode = proc.ExitCode;

            bool ok = exitCode == 0;
            onOutput(ok
                ? (uninstall ? "\n✓ Agent uninstalled!" : "\n✓ Agent installed successfully!")
                : $"\n✗ Installer exited with code {exitCode}");
            return ok;
        }

        private static void FlushLog(string logPath, Action<string> onOutput)
        {
            try
            {
                if (!File.Exists(logPath)) return;
                // Open with ReadWrite share so the elevated writer isn't blocked
                using var fs = new FileStream(logPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var sr = new StreamReader(fs, Encoding.UTF8);
                var lines = new List<string>();
                string? line;
                while ((line = sr.ReadLine()) != null) lines.Add(line);

                for (int i = _lastFlushedLine; i < lines.Count; i++)
                {
                    string trimmed = lines[i].TrimEnd();
                    if (!string.IsNullOrWhiteSpace(trimmed))
                        onOutput(trimmed);
                }
                _lastFlushedLine = lines.Count;
            }
            catch { }
        }
    }
}