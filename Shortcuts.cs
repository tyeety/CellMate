using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace CellMate
{
    public sealed class ShortcutBinding
    {
        public string Command { get; set; }
        public int Keys { get; set; }
    }

    public sealed class Preferences
    {
        public List<ShortcutBinding> Shortcuts { get; set; } = new List<ShortcutBinding>();

        public static Preferences Defaults() => new Preferences
        {
            Shortcuts = new List<ShortcutBinding>
            {
                new ShortcutBinding { Command = "Duplicate", Keys = (int)(System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.D) },
                new ShortcutBinding { Command = "PasteValues", Keys = (int)(System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift | System.Windows.Forms.Keys.V) },
                new ShortcutBinding { Command = "AutoFit", Keys = (int)(System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.O) },
                new ShortcutBinding { Command = "Settings", Keys = (int)(System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.K) }
            }
        };

        public Preferences Clone() => new Preferences { Shortcuts = Shortcuts.Select(s => new ShortcutBinding { Command = s.Command, Keys = s.Keys }).ToList() };

        public void Validate()
        {
            var expected = Defaults().Shortcuts.Select(s => s.Command).OrderBy(s => s).ToArray();
            if (Shortcuts == null || Shortcuts.Any(s => s == null) || !Shortcuts.Select(s => s.Command).OrderBy(s => s).SequenceEqual(expected))
                throw new InvalidDataException("The shortcut settings file is invalid.");
            foreach (var binding in Shortcuts)
                if (binding.Keys != 0) Shortcut.ToExcel((Keys)binding.Keys);
            if (Shortcuts.Where(s => s.Keys != 0).GroupBy(s => s.Keys).Any(g => g.Count() > 1))
                throw new InvalidDataException("Each shortcut can be assigned to only one command.");
        }
    }

    internal static class Shortcut
    {
        public static string ToExcel(Keys keys)
        {
            var code = keys & Keys.KeyCode;
            var modifiers = keys & Keys.Modifiers;
            if ((modifiers & (Keys.Control | Keys.Alt)) == 0 || (modifiers & ~(Keys.Control | Keys.Shift | Keys.Alt)) != 0)
                throw new InvalidDataException("A shortcut must include Ctrl or Alt.");
            string key;
            if (code >= Keys.A && code <= Keys.Z) key = code.ToString().ToLowerInvariant();
            else if (code >= Keys.D0 && code <= Keys.D9) key = ((int)code - (int)Keys.D0).ToString();
            else if (code >= Keys.F1 && code <= Keys.F12) key = "{" + code + "}";
            else throw new InvalidDataException("Use a letter A-Z, a number, or F1-F12 for the shortcut.");
            // These combinations close Excel/windows or invoke operating-system commands.
            if ((code == Keys.F4 && (modifiers & (Keys.Alt | Keys.Control)) != 0) ||
                (code == Keys.L && modifiers == Keys.Control) ||
                (code == Keys.F12 && modifiers == Keys.Control))
                throw new InvalidDataException("This key combination is reserved and cannot be used.");
            return ((keys & Keys.Control) != 0 ? "^" : "") + ((keys & Keys.Shift) != 0 ? "+" : "") + ((keys & Keys.Alt) != 0 ? "%" : "") + key;
        }

        public static string Display(int keys) => keys == 0 ? "Disabled" : new KeysConverter().ConvertToInvariantString((Keys)keys);
    }

    internal static class PreferenceStore
    {
        public static string FilePath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "CellMate", "settings.xml");
        public static Preferences Load()
        {
            if (!File.Exists(FilePath)) return Preferences.Defaults();
            using (var input = File.OpenRead(FilePath))
            {
                var preferences = (Preferences)new XmlSerializer(typeof(Preferences)).Deserialize(input);
                preferences.Validate();
                return preferences;
            }
        }

        public static void Save(Preferences preferences)
        {
            preferences.Validate();
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath));
            var pending = FilePath + ".pending";
            try
            {
                using (var output = File.Create(pending)) new XmlSerializer(typeof(Preferences)).Serialize(output, preferences);
                if (File.Exists(FilePath)) File.Replace(pending, FilePath, null);
                else File.Move(pending, FilePath);
            }
            finally { if (File.Exists(pending)) File.Delete(pending); }
        }
    }
}
