using System;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace FocusShield
{
    // Generates tray icons programmatically — no embedded resources needed.
    internal static class IconRenderer
    {
        // Green shield  = protection active
        // Amber shield  = just blocked something
        // Grey shield   = paused

        public static Icon CreateActiveIcon()  => BuildShield(Color.FromArgb(0, 185, 80),  Color.FromArgb(0, 120, 50));
        public static Icon CreateBlockedIcon() => BuildShield(Color.FromArgb(255, 190, 0),  Color.FromArgb(180, 120, 0));
        public static Icon CreatePausedIcon()  => BuildShield(Color.FromArgb(130, 130, 130), Color.FromArgb(80, 80, 80));

        private static Icon BuildShield(Color fill, Color shadow)
        {
            // 16x16 shield polygon
            var pts = new PointF[]
            {
                new(2f,  1f),
                new(14f, 1f),
                new(14f, 9f),
                new(8f,  15f),
                new(2f,  9f),
            };

            using var bmp = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);

                using var fillBrush    = new SolidBrush(fill);
                using var borderPen    = new Pen(shadow, 1.2f);
                using var highlightPen = new Pen(Color.FromArgb(120, 255, 255, 255), 0.8f);

                g.FillPolygon(fillBrush, pts);
                g.DrawPolygon(borderPen, pts);

                // Small highlight line near top-left for a slight 3-D feel
                g.DrawLine(highlightPen, 3f, 2f, 7f, 2f);
            }

            IntPtr hIcon = bmp.GetHicon();
            var icon = (Icon)Icon.FromHandle(hIcon).Clone(); // clone so we own the data
            DestroyIcon(hIcon);
            return icon;
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool DestroyIcon(IntPtr hIcon);
    }
}
