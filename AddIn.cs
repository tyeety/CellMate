using System;
using System.Collections.Generic;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace CellMate
{
    public sealed class AddIn : IExcelAddIn
    {
        internal static Preferences Preferences = Preferences.Defaults();
        private static readonly List<string> BoundKeys = new List<string>();
        private static EditShortcut editShortcut;
        internal static dynamic Application => ExcelDnaUtil.Application;

        public void AutoOpen()
        {
            string warning = null;
            try { Preferences = PreferenceStore.Load(); }
            catch (Exception ex) { warning = "Could not read settings. Default shortcuts will be used.\n" + ex.Message; }
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    Bind(Preferences);
                    editShortcut?.Dispose();
                    editShortcut = new EditShortcut();
                    editShortcut.Configure(Preferences);
                    if (warning != null) ShowError(warning);
                }
                catch (Exception ex) { ShowError("Could not activate shortcuts.\n" + ex.Message); }
            });
        }

        public void AutoClose()
        {
            editShortcut?.Dispose();
            editShortcut = null;
            try { Unbind(); } catch { /* Excel may already be shutting down. */ }
            Commands.ClearUndo();
        }

        private static void Unbind()
        {
            foreach (var key in BoundKeys) Application.OnKey(key, Type.Missing);
            BoundKeys.Clear();
        }

        private static void Bind(Preferences preferences)
        {
            Unbind();
            foreach (var binding in preferences.Shortcuts)
            {
                if (binding.Keys == 0) continue;
                var key = Shortcut.ToExcel((Keys)binding.Keys);
                Application.OnKey(key, "CellMate_" + binding.Command);
                BoundKeys.Add(key);
            }
            editShortcut?.Configure(preferences);
        }

        internal static void SavePreferences(Preferences next)
        {
            next.Validate();
            var previous = Preferences;
            try { Bind(next); PreferenceStore.Save(next); Preferences = next; }
            catch { Bind(previous); throw; }
        }

        internal static void ShowError(string message) => MessageBox.Show(new ExcelWindow(), message, "CellMate", MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }

    internal sealed class ExcelWindow : IWin32Window
    {
        public IntPtr Handle => ExcelDnaUtil.WindowHandle;
    }
}
