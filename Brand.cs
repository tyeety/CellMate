using System.Drawing;
using System.Drawing.Drawing2D;

namespace CellMate
{
    internal static class Brand
    {
        // Code-native artwork scales cleanly for the ribbon and settings window.
        public static Bitmap Icon(int size)
        {
            var bitmap = new Bitmap(size, size);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.ScaleTransform(size / 64f, size / 64f);
                using (var bg = new LinearGradientBrush(new Rectangle(0, 0, 64, 64), Theme.Primary, Theme.GradientEnd, 45))
                using (var row = new SolidBrush(Color.White))
                using (var copy = new SolidBrush(Color.FromArgb(231, 180, 255)))
                using (var plus = new Pen(Theme.Primary, 3.5f))
                {
                    g.FillEllipse(bg, 1, 1, 62, 62);
                    g.FillRectangle(row, 14, 17, 31, 10);
                    g.FillRectangle(copy, 19, 33, 31, 10);
                    g.DrawLine(plus, 35, 35, 35, 41);
                    g.DrawLine(plus, 32, 38, 38, 38);
                    using (var arrow = new Pen(Color.White, 2.5f))
                    {
                        g.DrawLines(arrow, new[] { new PointF(12, 31), new PointF(12, 43), new PointF(16, 43) });
                    }
                }
            }
            return bitmap;
        }
    }
}
