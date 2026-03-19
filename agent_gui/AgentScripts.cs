using System.IO;
using System.Reflection;
using System.Text;

namespace SnipeAgent
{
    static class AgentScripts
    {
        private const string AgentFile     = "DiskHealthAgent.ps1";
        private const string TrayFile      = "DiskHealthTray.ps1";
        private const string InstallerFile = "InstallDiskAgent.ps1";

        public static string Agent     => Load(AgentFile);
        public static string Tray      => Load(TrayFile);
        public static string Installer => Load(InstallerFile);

        public static void ExtractTo(string directory)
        {
            Directory.CreateDirectory(directory);
            var utf8bom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: true);
            File.WriteAllText(Path.Combine(directory, AgentFile),     Agent,     utf8bom);
            File.WriteAllText(Path.Combine(directory, TrayFile),      Tray,      utf8bom);
            File.WriteAllText(Path.Combine(directory, InstallerFile), Installer, utf8bom);
        }

        private static string Load(string fileName)
        {
            // 1. Embedded resource — always works in single-file publish
            var asm = Assembly.GetExecutingAssembly();
            foreach (string name in new[] { fileName, "SnipeAgent." + fileName })
            {
                using Stream? s = asm.GetManifestResourceStream(name);
                if (s is null) continue;
                using var r = new StreamReader(s, Encoding.UTF8);
                return r.ReadToEnd();
            }

            // 2. Loose file next to EXE (dev / debug builds)
            string base_ = AppContext.BaseDirectory;
            foreach (string path in new[]
            {
                Path.Combine(base_, fileName),
                Path.Combine(base_, "Scripts", fileName),
            })
                if (File.Exists(path)) return File.ReadAllText(path, Encoding.UTF8);

            throw new FileNotFoundException(
                $"Script '{fileName}' not found as embedded resource or loose file.", fileName);
        }
    }
}