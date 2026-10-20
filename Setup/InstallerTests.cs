using System;
using System.IO;
using System.Linq;
using Microsoft.Win32;

namespace CellMate.Setup
{
    // Runs against disposable paths and registry keys, never the real Excel installation.
    public static class InstallerTests
    {
        public static string Run()
        {
            int checks = 0;
            Action<bool, string> assert = (condition, label) => { if (!condition) throw new Exception(label); checks++; };
            string root = Path.Combine(Path.GetTempPath(), "CellMate-setup-tests-" + Guid.NewGuid().ToString("N"));
            string registry = @"Software\CellMate\SetupTests\" + Guid.NewGuid().ToString("N");
            var paths = new InstallPaths
            {
                Root = Path.Combine(root, "install"), LegacyRoot = Path.Combine(root, "legacy"),
                Settings = Path.Combine(root, @"preferences\settings.xml"), LegacySettings = Path.Combine(root, @"legacy-preferences\settings.xml"),
                OptionsKey = registry + @"\Options", ManagerKey = registry + @"\Manager", UninstallKey = registry + @"\Uninstall"
            };
            try
            {
                assert(ExcelInstallation.ReadMachine(Installer.Payload(true)) == 0x8664, "64-bit payload");
                assert(ExcelInstallation.ReadMachine(Installer.Payload(false)) == 0x14c, "32-bit payload");
                Directory.CreateDirectory(Path.GetDirectoryName(paths.LegacySettings));
                File.WriteAllText(paths.LegacySettings, "<saved-custom-shortcuts />");
                Directory.CreateDirectory(paths.LegacyRoot);
                File.WriteAllText(paths.LegacyBinary, "old add-in");
                using (var key = Registry.CurrentUser.CreateSubKey(paths.OptionsKey))
                {
                    key.SetValue("OPEN", "/R \"C:\\OtherAddin.xll\"");
                    key.SetValue("OPEN2", "/R \"" + paths.LegacyBinary + "\"");
                    key.SetValue("OPEN4", "/R \"" + paths.OldBinary + "\"");
                    key.SetValue("UserPreference", 71, RegistryValueKind.DWord);
                    key.SetValue("ExpandedPreference", "%TEMP%", RegistryValueKind.ExpandString);
                }
                using (var key = Registry.CurrentUser.CreateSubKey(paths.ManagerKey)) { key.SetValue(paths.LegacyBinary, ""); key.SetValue(@"C:\OtherAddin.xll", ""); }
                var installer = new Installer(paths, () => false);
                var excel = new ExcelInstallation { Path = "test-excel.exe", Is64Bit = true };
                installer.Install(excel);
                assert(Installer.SameBytes(File.ReadAllBytes(paths.Binary), Installer.Payload(true)), "installed payload verification");
                assert(File.ReadAllText(Path.Combine(paths.Root, "LICENSE")).Contains("Source Viewing License"), "license accompanies installed binaries");
                assert(File.ReadAllText(Path.Combine(paths.Root, "THIRD_PARTY_NOTICES.md")).Contains("Govert van Drimmelen"), "third-party notice accompanies installed binaries");
                assert(File.ReadAllText(paths.Settings) == "<saved-custom-shortcuts />", "preference migration");
                assert(!File.Exists(paths.LegacyBinary), "obsolete binary removed");
                assert(!File.Exists(paths.LegacySettings), "duplicate legacy preferences removed");
                using (var key = Registry.CurrentUser.OpenSubKey(paths.OptionsKey))
                {
                    assert((int)key.GetValue("UserPreference") == 71, "unrelated Excel option preserved");
                    assert(key.GetValueKind("ExpandedPreference") == RegistryValueKind.ExpandString, "registry value type preserved");
                    assert((string)key.GetValue("OPEN") == "/R \"C:\\OtherAddin.xll\"", "unrelated startup preserved");
                    assert((string)key.GetValue("OPEN2") == "/R \"" + paths.Binary + "\"", "existing startup slot reused");
                    assert(key.GetValue("OPEN4") == null, "duplicate registration removed");
                }
                using (var key = Registry.CurrentUser.OpenSubKey(paths.ManagerKey))
                { assert(key.GetValue(paths.LegacyBinary) == null, "old add-in manager entry removed"); assert(key.GetValue(@"C:\OtherAddin.xll") != null, "unrelated manager entry preserved"); }
                using (var key = Registry.CurrentUser.OpenSubKey(paths.UninstallKey))
                { assert((string)key.GetValue("DisplayVersion") == Theme.Version, "Apps registration"); assert(((string)key.GetValue("UninstallString")).EndsWith(" --uninstall"), "uninstaller command"); }
                File.WriteAllText(paths.Settings, "<my-new-custom-shortcuts />");
                installer.Install(excel);
                assert(File.ReadAllText(paths.Settings) == "<my-new-custom-shortcuts />", "repair preserves current preferences");
                using (var key = Registry.CurrentUser.OpenSubKey(paths.OptionsKey)) assert(key.GetValueNames().Count(n => Installer.IsOpenName(n) && paths.OwnsStartup(key.GetValue(n) as string)) == 1, "idempotent registration");
                File.WriteAllText(paths.Binary, "previous-version-payload");
                using (var key = Registry.CurrentUser.OpenSubKey(paths.UninstallKey, true)) key.SetValue("DisplayVersion", "old-version");
                installer.BeforeCommit = () => { throw new IOException("Injected installation failure"); };
                bool failed = false;
                try { installer.Install(excel); } catch (IOException) { failed = true; }
                assert(failed, "failure surfaced");
                assert(File.ReadAllText(paths.Binary) == "previous-version-payload", "rollback restores previous payload");
                using (var key = Registry.CurrentUser.OpenSubKey(paths.UninstallKey)) assert((string)key.GetValue("DisplayVersion") == "old-version", "rollback restores registration");
                assert(File.ReadAllText(paths.Settings) == "<my-new-custom-shortcuts />", "rollback preserves preferences");
                assert(!Directory.GetFiles(root, "*.pending", SearchOption.AllDirectories).Any(), "no temporary payloads left behind");
                failed = false;
                try { new Installer(paths, () => true).Install(excel); } catch (InvalidOperationException) { failed = true; }
                assert(failed && File.ReadAllText(paths.Binary) == "previous-version-payload", "open Excel blocks mutations");
                assert(!paths.OwnsBinary(Path.Combine(root, "elsewhere", "CellMate.xll")), "unrelated file excluded");
                File.WriteAllText(Path.Combine(paths.Root, "user-file.txt"), "keep");
                installer.Uninstall();
                assert(!File.Exists(paths.Binary), "uninstall removes payload");
                assert(!File.Exists(paths.Uninstaller), "uninstall removes setup copy");
                assert(File.ReadAllText(paths.Settings) == "<my-new-custom-shortcuts />", "uninstall preserves settings");
                assert(File.Exists(Path.Combine(paths.Root, "user-file.txt")), "uninstall preserves unrelated files");
                using (var key = Registry.CurrentUser.OpenSubKey(paths.UninstallKey)) assert(key == null, "uninstall removes Apps registration");
                using (var key = Registry.CurrentUser.OpenSubKey(paths.OptionsKey))
                { assert(key.GetValue("OPEN2") == null, "uninstall removes own startup"); assert(key.GetValue("OPEN") != null && (int)key.GetValue("UserPreference") == 71, "uninstall preserves unrelated options"); }
                return "PASS: " + checks + " installer checks (payloads, migration, preservation, repair, rollback, open Excel, uninstall).";
            }
            finally
            {
                Registry.CurrentUser.DeleteSubKeyTree(registry, false);
                string full = Path.GetFullPath(root);
                string expectedPrefix = Path.Combine(Path.GetFullPath(Path.GetTempPath()), "CellMate-setup-tests-");
                if (full.StartsWith(expectedPrefix, StringComparison.OrdinalIgnoreCase) && Directory.Exists(full)) Directory.Delete(full, true);
            }
        }
    }
}
