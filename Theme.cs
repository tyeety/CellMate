using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace CellMate
{
    // Colors and corner radii from shatootsoft.com, verified against its live CSS.
    internal static class Theme
    {
        public static readonly Color Primary = ColorTranslator.FromHtml("#7400A8");
        public static readonly Color GradientEnd = ColorTranslator.FromHtml("#8F32B9");
        public static readonly Color Hover = ColorTranslator.FromHtml("#9F38D7");
        public static readonly Color Background = ColorTranslator.FromHtml("#FFF7FC");
        public static readonly Color Text = ColorTranslator.FromHtml("#201923");
        public static readonly Color Muted = ColorTranslator.FromHtml("#7A4B93");
        public static readonly Color Body = ColorTranslator.FromHtml("#4E4353");
        public static readonly Color Border = ColorTranslator.FromHtml("#D2C1D5");
        public static readonly Color Tint = ColorTranslator.FromHtml("#F4E4F5");
        public static readonly Color Error = ColorTranslator.FromHtml("#BA1A1A");
        public const string Version = "0.2.0";

        public static GraphicsPath Round(RectangleF rect, float radius)
        {
            var path = new GraphicsPath();
            float d = Math.Min(radius * 2, Math.Min(rect.Width, rect.Height));
            path.AddArc(rect.X, rect.Y, d, d, 180, 90);
            path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
            path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
            path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        public static Label Label(string text, float size = 10, bool bold = false, Color? color = null) => new Label
        {
            Text = text, AutoSize = false, ForeColor = color ?? Text, BackColor = Color.Transparent,
            Font = new Font("Segoe UI", size, bold ? FontStyle.Bold : FontStyle.Regular),
            TextAlign = ContentAlignment.MiddleLeft, UseMnemonic = false
        };

        public static void StyleWindow(Form form)
        {
            form.BackColor = Background;
            form.ForeColor = Text;
            form.Font = new Font("Segoe UI", 10);
            form.AutoScaleMode = AutoScaleMode.Dpi;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.FormBorderStyle = FormBorderStyle.FixedSingle;
            form.MaximizeBox = false;
            using (var stream = typeof(Theme).Assembly.GetManifestResourceStream("CellMate.ico"))
                if (stream != null) using (var icon = new Icon(stream)) form.Icon = (Icon)icon.Clone();
        }
    }

    internal sealed class BrandButton : Button
    {
        public bool Primary { get; set; }
        private bool hovered;
        public BrandButton()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            Cursor = Cursors.Hand;
            Height = 44;
            Font = new Font("Segoe UI", 10, FontStyle.Bold);
            BackColor = Theme.Background;
            UseVisualStyleBackColor = false;
        }
        protected override void OnMouseEnter(EventArgs e) { hovered = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { hovered = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            e.Graphics.Clear(Parent?.BackColor ?? Theme.Background);
            using (var path = Theme.Round(new RectangleF(1, 1, Width - 3, Height - 3), 8))
            {
                if (Primary)
                {
                    using (var brush = new LinearGradientBrush(ClientRectangle, !Enabled ? Theme.Border : hovered ? Theme.Hover : Theme.Primary,
                        !Enabled ? Theme.Border : hovered ? Theme.Hover : Theme.GradientEnd, LinearGradientMode.Vertical)) e.Graphics.FillPath(brush, path);
                }
                else
                {
                    using (var fill = new SolidBrush(hovered ? Theme.Tint : Color.White)) e.Graphics.FillPath(fill, path);
                    using (var pen = new Pen(Enabled ? Theme.Border : Theme.Tint)) e.Graphics.DrawPath(pen, path);
                }
            }
            TextRenderer.DrawText(e.Graphics, Text, Font, ClientRectangle, Primary ? Color.White : Enabled ? Theme.Primary : Theme.Border,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
            if (Focused && ShowFocusCues) ControlPaint.DrawFocusRectangle(e.Graphics, new Rectangle(7, 7, Width - 15, Height - 15));
        }
    }

    internal class Card : Panel
    {
        public Card() { DoubleBuffered = true; BackColor = Theme.Background; Padding = new Padding(20); }
        protected override void OnPaintBackground(PaintEventArgs e)
        {
            e.Graphics.Clear(Parent?.BackColor ?? Theme.Background);
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (var path = Theme.Round(new RectangleF(1, 1, Width - 3, Height - 3), 16))
            using (var brush = new SolidBrush(Color.White))
            using (var pen = new Pen(Theme.Border))
            { e.Graphics.FillPath(brush, path); e.Graphics.DrawPath(pen, path); }
        }
    }

    internal sealed class BrandRail : Panel
    {
        private readonly Bitmap logo = Brand.Icon(64);
        public BrandRail() { DoubleBuffered = true; }
        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var fill = new LinearGradientBrush(ClientRectangle, Theme.Primary, Color.FromArgb(65, 20, 91), LinearGradientMode.ForwardDiagonal)) g.FillRectangle(fill, ClientRectangle);
            using (var glow = new SolidBrush(Color.FromArgb(25, 255, 255, 255)))
            { g.FillEllipse(glow, Width - 180, -70, 300, 300); g.FillEllipse(glow, -160, Height - 280, 330, 330); }
            g.DrawImage(logo, 28, 34, 64, 64);
            Draw(g, "CellMate", 25, FontStyle.Bold, Color.White, 28, 108, Width - 50, 45);
            Draw(g, "A little help.\nA lot less work.", 17, FontStyle.Bold, Color.White, 30, 176, Width - 48, 80);
            Draw(g, "Your everyday Excel tasks,\none shortcut away.", 10, FontStyle.Regular, Color.FromArgb(235, 210, 246), 30, 264, Width - 48, 60);
            // A small spreadsheet illustration drawn as vector artwork.
            float x = 30, y = Height - 210, w = Width - 60;
            using (var path = Theme.Round(new RectangleF(x, y, w, 104), 12))
            using (var fill = new SolidBrush(Color.FromArgb(28, 255, 255, 255))) g.FillPath(fill, path);
            using (var line = new Pen(Color.FromArgb(55, 255, 255, 255)))
            {
                for (int row = 1; row < 4; row++) g.DrawLine(line, x + 12, y + row * 25, x + w - 12, y + row * 25);
                for (int col = 1; col < 4; col++) g.DrawLine(line, x + col * w / 4, y + 10, x + col * w / 4, y + 94);
            }
            using (var stroke = new Pen(Color.FromArgb(231, 180, 255), 2)) g.DrawRectangle(stroke, x + 8, y + 28, w - 16, 24);
            Draw(g, "SHATOOTSOFT", 9, FontStyle.Bold, Color.FromArgb(231, 180, 255), 30, Height - 70, Width - 50, 25);
            Draw(g, "shatootsoft.com", 9, FontStyle.Regular, Color.FromArgb(210, 182, 223), 30, Height - 47, Width - 50, 22);
        }
        private static void Draw(Graphics g, string text, float size, FontStyle style, Color color, int x, int y, int w, int h)
        {
            using (var font = new Font("Segoe UI", size, style))
            using (var brush = new SolidBrush(color)) g.DrawString(text, font, brush, new RectangleF(x, y, w, h));
        }
        protected override void Dispose(bool disposing) { if (disposing) logo.Dispose(); base.Dispose(disposing); }
    }
}
