using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Text;

namespace FocusShield
{
    // Generates tray icons programmatically — no embedded resources needed.
    internal static class IconRenderer
    {
        // Green shield  = protection active
        // Amber shield  = just kept something in the background
        // Grey shield   = paused

        public static Icon CreateActiveIcon()  => BuildShield(Color.FromArgb(0, 185, 80),  Color.FromArgb(0, 120, 50));
        public static Icon CreateBlockedIcon() => BuildShield(Color.FromArgb(255, 190, 0),  Color.FromArgb(180, 120, 0));
        public static Icon CreatePausedIcon()  => BuildShield(Color.FromArgb(130, 130, 130), Color.FromArgb(80, 80, 80));

        private const int IconSize = 16;

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

            using var bmp = new Bitmap(IconSize, IconSize, PixelFormat.Format32bppArgb);
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

            return ToIcon(bmp);
        }

        /// <summary>
        /// Packs the bitmap into an in-memory .ico and builds the Icon from that.
        /// Bitmap.GetHicon would hand back an Icon that only borrows an HICON we
        /// then have to free ourselves — and both cloning and disposing that Icon
        /// leave the tray pointing at a freed handle.
        /// </summary>
        private static Icon ToIcon(Bitmap bmp)
        {
            int width  = bmp.Width;
            int height = bmp.Height;

            int xorSize = width * 4 * height;                  // 32bpp colour data
            int andSize = (width + 31) / 32 * 4 * height;       // 1bpp mask, rows padded to 4 bytes

            using var stream = new MemoryStream();
            using (var w = new BinaryWriter(stream, Encoding.Unicode, leaveOpen: true))
            {
                // ICONDIR
                w.Write((short)0);                  // reserved
                w.Write((short)1);                  // type: icon
                w.Write((short)1);                  // image count

                // ICONDIRENTRY
                w.Write((byte)width);
                w.Write((byte)height);
                w.Write((byte)0);                   // palette entries
                w.Write((byte)0);                   // reserved
                w.Write((short)1);                  // colour planes
                w.Write((short)32);                 // bits per pixel
                w.Write(40 + xorSize + andSize);    // size of the image data
                w.Write(22);                        // offset of the image data

                // BITMAPINFOHEADER — the height spans both the colour and mask bitmaps
                w.Write(40);
                w.Write(width);
                w.Write(height * 2);
                w.Write((short)1);
                w.Write((short)32);
                w.Write(0);                         // BI_RGB
                w.Write(xorSize + andSize);
                w.Write(0);                         // pixels per metre X
                w.Write(0);                         // pixels per metre Y
                w.Write(0);                         // palette colours used
                w.Write(0);                         // important colours

                // Colour data, bottom-up BGRA
                for (int y = height - 1; y >= 0; y--)
                {
                    for (int x = 0; x < width; x++)
                    {
                        Color c = bmp.GetPixel(x, y);
                        w.Write(c.B);
                        w.Write(c.G);
                        w.Write(c.R);
                        w.Write(c.A);
                    }
                }

                // AND mask left empty — transparency comes from the alpha channel
                w.Write(new byte[andSize]);
            }

            stream.Position = 0;
            return new Icon(stream);
        }
    }
}
