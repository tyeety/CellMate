using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using Microsoft.Win32;

namespace CellMate.Setup
{
    internal sealed class ExcelInstallation
    {
        public string Path;
        public bool Is64Bit;
        public string Description => "Excel for Windows  /  " + (Is64Bit ? "64-bit" : "32-bit");
        public static ExcelInstallation Find()
        {
            var candidates = new List<string>();
            foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
            foreach (var hive in new[] { RegistryHive.LocalMachine, RegistryHive.CurrentUser })
            using (var root = RegistryKey.OpenBaseKey(hive, view))
            {
                using (var key = root.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe"))
                    if (key?.GetValue("") is string path) candidates.Add(path.Trim('"'));
                using (var key = root.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"))
                    if (key?.GetValue("InstallationPath") is string office) candidates.Add(System.IO.Path.Combine(office, @"root\Office16\EXCEL.EXE"));
            }
            foreach (var dir in new[] { Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86) })
            foreach (var sub in new[] { @"Microsoft Office\root\Office16\EXCEL.EXE", @"Microsoft Office\Office16\EXCEL.EXE" })
                if (!string.IsNullOrEmpty(dir)) candidates.Add(System.IO.Path.Combine(dir, sub));
            foreach (string path in candidates.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                if (!File.Exists(path)) continue;
                if (FileVersionInfo.GetVersionInfo(path).FileMajorPart < 16) continue;
                ushort machine = ReadMachine(File.ReadAllBytes(path));
                if (machine == 0x8664 || machine == 0x14c) return new ExcelInstallation { Path = path, Is64Bit = machine == 0x8664 };
            }
            return null;
        }

        internal static ushort ReadMachine(byte[] bytes)
        {
            if (bytes.Length < 64 || bytes[0] != 'M' || bytes[1] != 'Z') throw new InvalidDataException("Invalid Windows executable.");
            int offset = BitConverter.ToInt32(bytes, 0x3c);
            if (offset < 0 || offset > bytes.Length - 6 || BitConverter.ToInt32(bytes, offset) != 0x4550) throw new InvalidDataException("Invalid PE header.");
            return BitConverter.ToUInt16(bytes, offset + 4);
        }
    }

    internal sealed class InstallPaths
    {
        public string Root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CellMate");
        public string LegacyRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "RowMate");
        public string Settings = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"CellMate\settings.xml");
        public string LegacySettings = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"RowMate\settings.xml");
        public string OptionsKey = @"Software\Microsoft\Office\16.0\Excel\Options";
        public string ManagerKey = @"Software\Microsoft\Office\16.0\Excel\Add-in Manager";
        public string UninstallKey = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\CellMate";
        public string Binary => Path.Combine(Root, Theme.Version, "CellMate.xll");
        public string Uninstaller => Path.Combine(Root, "CellMateSetup.exe");
        public string OldBinary => Path.Combine(Root, "CellMate.xll");
        public string LegacyBinary => Path.Combine(LegacyRoot, "RowMate.xll");
        public bool OwnsBinary(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return false;
            string full;
            try { full = Path.GetFullPath(path); } catch { return false; }
            if (string.Equals(full, LegacyBinary, StringComparison.OrdinalIgnoreCase) || string.Equals(full, OldBinary, StringComparison.OrdinalIgnoreCase)) return true;
            string parent = Path.GetDirectoryName(full);
            return string.Equals(Path.GetFileName(full), "CellMate.xll", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(Path.GetDirectoryName(parent), Path.GetFullPath(Root), StringComparison.OrdinalIgnoreCase) &&
                Version.TryParse(Path.GetFileName(parent), out _);
        }
        public bool OwnsStartup(string value)
        {
            if (value == null || !value.StartsWith("/R ", StringComparison.OrdinalIgnoreCase)) return false;
            return OwnsBinary(value.Substring(3).Trim().Trim('"'));
        }
    }

    internal sealed class RegistryJournal
    {
        private readonly List<Action> undo = new List<Action>();
        public void Set(string path, string name, object value, RegistryValueKind kind = RegistryValueKind.String)
        {
            Remember(path, name);
            using (var key = Registry.CurrentUser.CreateSubKey(path)) key.SetValue(name, value, kind);
        }
        public void Delete(string path, string name)
        {
            Remember(path, name);
            using (var key = Registry.CurrentUser.OpenSubKey(path, true)) key?.DeleteValue(name, false);
        }
        private void Remember(string path, string name)
        {
            using (var key = Registry.CurrentUser.OpenSubKey(path))
            {
                bool exists = key != null && key.GetValueNames().Contains(name, StringComparer.OrdinalIgnoreCase);
                object previous = exists ? key.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames) : null;
                var kind = exists ? key.GetValueKind(name) : RegistryValueKind.String;
                undo.Add(() => { using (var restore = Registry.CurrentUser.CreateSubKey(path)) { if (exists) restore.SetValue(name, previous, kind); else restore.DeleteValue(name, false); } });
            }
        }
        public void Rollback() { foreach (var action in undo.AsEnumerable().Reverse()) action(); }
    }

    internal sealed class Installer
    {
        internal readonly InstallPaths Paths;
        private readonly Func<bool> excelRunning;
        internal Action BeforeCommit;
        public Installer(InstallPaths paths = null, Func<bool> running = null) { Paths = paths ?? new InstallPaths(); excelRunning = running ?? IsExcelRunning; }
        public static bool IsExcelRunning()
        {
            var processes = Process.GetProcessesByName("EXCEL");
            try { return processes.Length != 0; } finally { foreach (var process in processes) process.Dispose(); }
        }
        internal static byte[] Payload(bool is64Bit)
            => Resource(is64Bit ? "CellMate.x64" : "CellMate.x86");

        internal static byte[] Resource(string name)
        {
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(name))
            {
                if (stream == null) throw new InvalidDataException("The installer payload is missing. Download a fresh copy of CellMate Setup.");
                using (var memory = new MemoryStream()) { stream.CopyTo(memory); return memory.ToArray(); }
            }
        }
        internal static bool SameBytes(byte[] a, byte[] b)
        {
            using (var hash = SHA256.Create()) return hash.ComputeHash(a).SequenceEqual(hash.ComputeHash(b));
        }
        internal static void AtomicWrite(string path, byte[] bytes)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            string pending = path + "." + Guid.NewGuid().ToString("N") + ".pending";
            try
            {
                File.WriteAllBytes(pending, bytes);
                if (File.Exists(path)) File.Replace(pending, path, null);
                else File.Move(pending, path);
            }
            finally { if (File.Exists(pending)) File.Delete(pending); }
        }
        public void Install(ExcelInstallation excel)
        {
            if (excel == null) throw new InvalidOperationException("Install Excel 2016 or later for Windows, then run CellMate Setup again.");
            if (excelRunning()) throw new InvalidOperationException("Save your work and close Excel before installing CellMate.");
            var payload = Payload(excel.Is64Bit);
            if (ExcelInstallation.ReadMachine(payload) != (excel.Is64Bit ? 0x8664 : 0x14c)) throw new InvalidDataException("The add-in architecture does not match Excel.");
            var registry = new RegistryJournal();
            var files = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            Action<string, byte[]> write = (path, bytes) =>
            {
                if (File.Exists(path) && SameBytes(File.ReadAllBytes(path), bytes)) return;
                byte[] previous = File.Exists(path) ? File.ReadAllBytes(path) : null;
                AtomicWrite(path, bytes);
                files.Add(path, previous);
            };
            try
            {
                write(Paths.Binary, payload);
                write(Path.Combine(Paths.Root, "LICENSE"), Resource("CellMate.License"));
                write(Path.Combine(Paths.Root, "THIRD_PARTY_NOTICES.md"), Resource("CellMate.ThirdPartyNotices"));
                if (!string.Equals(Assembly.GetExecutingAssembly().Location, Paths.Uninstaller, StringComparison.OrdinalIgnoreCase))
                    write(Paths.Uninstaller, File.ReadAllBytes(Assembly.GetExecutingAssembly().Location));
                if (!File.Exists(Paths.Settings) && File.Exists(Paths.LegacySettings)) write(Paths.Settings, File.ReadAllBytes(Paths.LegacySettings));
                if (excelRunning()) throw new InvalidOperationException("Excel opened during installation. Close it and try again.");
                var owned = new List<string>();
                string slot = "OPEN";
                using (var key = Registry.CurrentUser.OpenSubKey(Paths.OptionsKey))
                {
                    var names = key?.GetValueNames() ?? new string[0];
                    foreach (string name in names)
                        if (IsOpenName(name) && Paths.OwnsStartup(key.GetValue(name) as string)) owned.Add(name);
                    if (owned.Count != 0) slot = owned[0];
                    else { int i = 0; while (names.Contains(slot, StringComparer.OrdinalIgnoreCase)) slot = "OPEN" + (++i); }
                }
                registry.Set(Paths.OptionsKey, slot, "/R \"" + Paths.Binary + "\"");
                foreach (var name in owned.Where(n => n != slot)) registry.Delete(Paths.OptionsKey, name);
                CleanManager(registry);
                registry.Set(Paths.UninstallKey, "DisplayName", "CellMate");
                registry.Set(Paths.UninstallKey, "DisplayVersion", Theme.Version);
                registry.Set(Paths.UninstallKey, "Publisher", "ShatootSoft");
                registry.Set(Paths.UninstallKey, "DisplayIcon", Paths.Uninstaller);
                registry.Set(Paths.UninstallKey, "InstallLocation", Paths.Root);
                registry.Set(Paths.UninstallKey, "UninstallString", "\"" + Paths.Uninstaller + "\" --uninstall");
                registry.Set(Paths.UninstallKey, "URLInfoAbout", "https://shatootsoft.com");
                registry.Set(Paths.UninstallKey, "EstimatedSize", (payload.Length + new FileInfo(Paths.Uninstaller).Length) / 1024, RegistryValueKind.DWord);
                registry.Set(Paths.UninstallKey, "NoModify", 1, RegistryValueKind.DWord);
                registry.Set(Paths.UninstallKey, "NoRepair", 1, RegistryValueKind.DWord);
                BeforeCommit?.Invoke();
                if (!SameBytes(File.ReadAllBytes(Paths.Binary), payload)) throw new IOException("The installed add-in did not pass verification.");
            }
            catch (Exception original)
            {
                var errors = new List<Exception> { original };
                try { registry.Rollback(); } catch (Exception ex) { errors.Add(ex); }
                foreach (var file in files.Reverse())
                    try { if (file.Value == null) File.Delete(file.Key); else AtomicWrite(file.Key, file.Value); } catch (Exception ex) { errors.Add(ex); }
                if (errors.Count > 1) throw new AggregateException("Installation failed and could not be completely rolled back.", errors);
                throw;
            }
            // Cleanup is deliberately limited to known, obsolete add-in files.
            foreach (var path in KnownBinaries().Where(p => !string.Equals(p, Paths.Binary, StringComparison.OrdinalIgnoreCase))) TryDelete(path);
            if (File.Exists(Paths.LegacySettings) && File.Exists(Paths.Settings) && SameBytes(File.ReadAllBytes(Paths.LegacySettings), File.ReadAllBytes(Paths.Settings))) TryDelete(Paths.LegacySettings);
            RemoveEmpty(Path.GetDirectoryName(Paths.LegacySettings));
            RemoveEmpty(Paths.LegacyRoot);
        }
        private void CleanManager(RegistryJournal journal)
        {
            using (var key = Registry.CurrentUser.OpenSubKey(Paths.ManagerKey))
                if (key != null) foreach (var name in key.GetValueNames()) if (Paths.OwnsBinary(name)) journal.Delete(Paths.ManagerKey, name);
        }
        internal static bool IsOpenName(string name) => name == "OPEN" || (name.StartsWith("OPEN", StringComparison.OrdinalIgnoreCase) && name.Length > 4 && name.Substring(4).All(char.IsDigit));
        private IEnumerable<string> KnownBinaries()
        {
            var result = new List<string> { Paths.OldBinary, Paths.LegacyBinary };
            if (Directory.Exists(Paths.Root)) foreach (var dir in Directory.GetDirectories(Paths.Root))
                if (Version.TryParse(Path.GetFileName(dir), out _)) result.Add(Path.Combine(dir, "CellMate.xll"));
            return result.Where(File.Exists);
        }
        public void Uninstall()
        {
            if (excelRunning()) throw new InvalidOperationException("Save your work and close Excel before removing CellMate.");
            // A failed deletion leaves the registration intact so removal can be retried.
            foreach (var path in KnownBinaries().ToArray()) File.Delete(path);
            using (var key = Registry.CurrentUser.OpenSubKey(Paths.OptionsKey, true))
                if (key != null) foreach (var name in key.GetValueNames()) if (IsOpenName(name) && Paths.OwnsStartup(key.GetValue(name) as string)) key.DeleteValue(name, false);
            CleanManager(new RegistryJournal());
            Registry.CurrentUser.DeleteSubKeyTree(Paths.UninstallKey, false);
            TryDelete(Path.Combine(Paths.Root, "LICENSE"));
            TryDelete(Path.Combine(Paths.Root, "THIRD_PARTY_NOTICES.md"));
            if (!string.Equals(Assembly.GetExecutingAssembly().Location, Paths.Uninstaller, StringComparison.OrdinalIgnoreCase)) TryDelete(Paths.Uninstaller);
            if (Directory.Exists(Paths.Root)) foreach (var dir in Directory.GetDirectories(Paths.Root)) if (Version.TryParse(Path.GetFileName(dir), out _)) RemoveEmpty(dir);
            RemoveEmpty(Paths.Root);
            RemoveEmpty(Paths.LegacyRoot);
        }
        internal static void TryDelete(string path) { try { if (File.Exists(path)) File.Delete(path); RemoveEmpty(Path.GetDirectoryName(path)); } catch (IOException) { } catch (UnauthorizedAccessException) { } }
        internal static void RemoveEmpty(string path) { if (Directory.Exists(path) && !Directory.EnumerateFileSystemEntries(path).Any()) Directory.Delete(path); }
        internal void DeleteUninstallerAfterExit()
        {
            if (!string.Equals(Assembly.GetExecutingAssembly().Location, Paths.Uninstaller, StringComparison.OrdinalIgnoreCase)) return;
            // The running executable cannot delete itself; a hidden Windows helper removes only this exact file after exit.
            string escaped = Paths.Uninstaller.Replace("'", "''");
            string root = Paths.Root.Replace("'", "''");
            string command = "Wait-Process -Id " + Process.GetCurrentProcess().Id + " -ErrorAction SilentlyContinue; Remove-Item -LiteralPath '" + escaped + "' -Force; if (!(Get-ChildItem -LiteralPath '" + root + "' -Force)) { Remove-Item -LiteralPath '" + root + "' }";
            Process.Start(new ProcessStartInfo("powershell.exe", "-NoProfile -NonInteractive -EncodedCommand " + Convert.ToBase64String(System.Text.Encoding.Unicode.GetBytes(command)))
                { UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden });
        }
    }
}
