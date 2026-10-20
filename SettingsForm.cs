using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace CellMate
{
    internal sealed class ShortcutBox : TextBox
    {
        public int ShortcutKeys { get; private set; }
        public ShortcutBox(int keys)
        {
            ReadOnly = true;
            RightToLeft = RightToLeft.No;
            TextAlign = HorizontalAlignment.Center;
            ShortcutsEnabled = false;
            BackColor = Color.White;
            ForeColor = Theme.Primary;
            Font = new Font("Segoe UI", 11, FontStyle.Bold);
            SetKeys(keys);
        }
        public void SetKeys(int keys) { ShortcutKeys = keys; Text = Shortcut.Display(keys); }
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if ((keyData & Keys.KeyCode) == Keys.Tab || keyData == Keys.Escape) return base.ProcessCmdKey(ref msg, keyData);
            if (keyData == Keys.Back || keyData == Keys.Delete) { SetKeys(0); return true; }
            var code = keyData & Keys.KeyCode;
            if (code == Keys.ControlKey || code == Keys.ShiftKey || code == Keys.Menu) return true;
            try { Shortcut.ToExcel(keyData); SetKeys((int)keyData); }
            catch (Exception ex) { AddIn.ShowError(ex.Message); }
            return true;
        }
    }

    internal sealed class SettingsForm : Form
    {
        private readonly Dictionary<string, ShortcutBox> boxes = new Dictionary<string, ShortcutBox>();
        private readonly Bitmap logo = Brand.Icon(56);
        public SettingsForm()
        {
            Theme.StyleWindow(this);
            Text = "CellMate Settings";
            ClientSize = new Size(760, 586);
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            RightToLeft = RightToLeft.No;
            RightToLeftLayout = false;
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(28, 22, 28, 18), ColumnCount = 1, RowCount = 7 };
            foreach (int height in new[] { 90, 44, 264, 52, 60, 34 }) layout.RowStyles.Add(new RowStyle(SizeType.Absolute, height));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var header = new Panel { Dock = DockStyle.Fill };
            header.Controls.Add(new PictureBox { Image = logo, SizeMode = PictureBoxSizeMode.CenterImage, Bounds = new Rectangle(0, 12, 60, 60) });
            header.Controls.Add(Position(Theme.Label("CELLMATE / PREFERENCES", 9, true, Theme.Muted), 78, 6, 560, 24));
            header.Controls.Add(Position(Theme.Label("Make Excel yours.", 23, true), 76, 28, 580, 52));
            layout.Controls.Add(header, 0, 0);
            var subtitle = Theme.Label("Pick your shortcuts. Keep your flow.", 11, false, Theme.Body);
            subtitle.Dock = DockStyle.Fill;
            layout.Controls.Add(subtitle, 0, 1);
            var card = new Card { Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 8), Padding = new Padding(22, 14, 22, 14) };
            var grid = new TableLayoutPanel { Dock = DockStyle.Fill, BackColor = Color.White, ColumnCount = 2, RowCount = 4 };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 65));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
            var names = new Dictionary<string, string[]>
            {
                ["Duplicate"] = new[] { "Duplicate selected rows", "Copy full rows with formulas and formatting." },
                ["PasteValues"] = new[] { "Paste values only", "Keep your destination formatting intact." },
                ["AutoFit"] = new[] { "AutoFit columns", "Give selected columns the space they need." },
                ["Settings"] = new[] { "Open settings", "Your shortcuts, always within reach." }
            };
            int row = 0;
            foreach (var binding in AddIn.Preferences.Shortcuts)
            {
                grid.RowStyles.Add(new RowStyle(SizeType.Percent, 25));
                var name = names[binding.Command];
                var description = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };
                description.Controls.Add(Position(Theme.Label(name[0], 10, true), 0, 0, 385, 27));
                description.Controls.Add(Position(Theme.Label(name[1], 9, false, Theme.Body), 0, 26, 385, 22));
                grid.Controls.Add(description, 0, row);
                var box = new ShortcutBox(binding.Keys) { Dock = DockStyle.Fill, Margin = new Padding(0, 13, 0, 12), AccessibleName = name[0] };
                boxes.Add(binding.Command, box);
                grid.Controls.Add(box, 1, row++);
            }
            card.Controls.Add(grid);
            layout.Controls.Add(card, 0, 2);
            var hint = Theme.Label("Click a box and press Ctrl or Alt with a key. Delete disables the shortcut.\nActive shortcuts replace Excel defaults, including Ctrl+D (Fill Down).", 9, false, Theme.Body);
            hint.Dock = DockStyle.Fill;
            layout.Controls.Add(hint, 0, 3);
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(0, 6, 0, 0) };
            var save = new BrandButton { Text = "Save changes", Primary = true, Width = 152 };
            save.Click += (_, __) => Save();
            var cancel = new BrandButton { Text = "Cancel", Width = 100, DialogResult = DialogResult.Cancel };
            var reset = new BrandButton { Text = "Restore defaults", Width = 155 };
            reset.Click += (_, __) => { foreach (var binding in Preferences.Defaults().Shortcuts) boxes[binding.Command].SetKeys(binding.Keys); };
            buttons.Controls.AddRange(new Control[] { save, cancel, reset });
            layout.Controls.Add(buttons, 0, 4);
            var footer = Theme.Label("SHATOOTSOFT    /    CellMate " + Theme.Version, 9, false, Theme.Muted);
            footer.Dock = DockStyle.Fill;
            layout.Controls.Add(footer, 0, 5);
            Controls.Add(layout);
            AcceptButton = save;
            CancelButton = cancel;
        }
        private static T Position<T>(T control, int x, int y, int w, int h) where T : Control { control.Bounds = new Rectangle(x, y, w, h); return control; }
        private void Save()
        {
            try
            {
                var next = new Preferences { Shortcuts = boxes.Select(pair => new ShortcutBinding { Command = pair.Key, Keys = pair.Value.ShortcutKeys }).ToList() };
                AddIn.SavePreferences(next);
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex) { AddIn.ShowError(ex.Message); }
        }
        protected override void Dispose(bool disposing) { base.Dispose(disposing); if (disposing) logo.Dispose(); }
    }
}
