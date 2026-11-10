using System;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace VisioSizeTools
{
    /// <summary>
    /// Draws the ribbon icons at runtime so they are crisp at any DPI and need no image resources.
    /// The large menu button is authored on a 32x32 grid; menu items on a 16x16 grid (like Align/Position).
    /// Convention: solid blue = the reference shape, light = a shape being resized, dashed = its new extent.
    /// </summary>
    internal static class Icons
    {
        private static readonly Color Accent = Color.FromArgb(57, 85, 163);       // Visio blue
        private static readonly Color AccentEdge = Color.FromArgb(38, 58, 115);
        private static readonly Color AccentLight = Color.FromArgb(205, 216, 240);
        private static readonly Color Outline = Color.FromArgb(96, 96, 96);
        private static readonly Color Guide = Color.FromArgb(214, 82, 34);        // orange: target extent

        public static Bitmap Create(string controlId)
        {
            bool large = controlId == "vstSizeMenu";
            int grid = large ? 32 : 16;
            int px = PixelSize(grid);

            var bmp = new Bitmap(px, px, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.Transparent);
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.ScaleTransform(px / (float)grid, px / (float)grid);

                switch (controlId)
                {
                    case "vstSizeTallest": DrawTallest(g); break;
                    case "vstSizeShortest": DrawShortest(g); break;
                    case "vstSizeWidest": DrawWidest(g); break;
                    case "vstSizeNarrowest": DrawNarrowest(g); break;
                    case "vstSizeBothLarger": DrawBothLarger(g); break;
                    case "vstSizeBothSmaller": DrawBothSmaller(g); break;
                    case "vstSizeFirstHeight": DrawFirstHeight(g); break;
                    case "vstSizeFirstWidth": DrawFirstWidth(g); break;
                    default: DrawMenu(g); break;
                }
            }
            return bmp;
        }

        private static int PixelSize(int logical)
        {
            try
            {
                using (Graphics screen = Graphics.FromHwnd(IntPtr.Zero))
                    return Math.Max(logical, (int)Math.Round(logical * screen.DpiX / 96f));
            }
            catch
            {
                return logical;
            }
        }

        // ---- 16x16 menu item icons -------------------------------------------------------

        private static void DrawTallest(Graphics g)
        {
            Solid(g, 2, 1, 5, 14);
            Light(g, 9, 8, 5, 7);
            Dashed(g, 9, 1, 5, 14);
        }

        private static void DrawShortest(Graphics g)
        {
            Solid(g, 2, 9, 5, 6);
            Light(g, 9, 1, 5, 14);
            HLine(g, 0, 16, 9);
        }

        private static void DrawWidest(Graphics g)
        {
            Solid(g, 1, 2, 14, 5);
            Light(g, 1, 9, 7, 5);
            Dashed(g, 1, 9, 14, 5);
        }

        private static void DrawNarrowest(Graphics g)
        {
            Solid(g, 1, 2, 6, 5);
            Light(g, 1, 9, 14, 5);
            VLine(g, 7, 0, 16);
        }

        private static void DrawBothLarger(Graphics g)
        {
            Light(g, 1, 1, 7, 7);
            Dashed(g, 1, 1, 14, 14);
            Solid(g, 10, 10, 5, 5); // corner of the largest extent
        }

        private static void DrawBothSmaller(Graphics g)
        {
            Light(g, 1, 1, 14, 14);
            Solid(g, 1, 1, 7, 7);
            using (var p = new Pen(Guide, 1f) { DashPattern = new[] { 2f, 1f } })
            {
                g.DrawLine(p, 8.5f, 1, 8.5f, 9);
                g.DrawLine(p, 1, 8.5f, 9, 8.5f);
            }
        }

        // The reference is the first shape, with one shape larger and one smaller than it.
        private static void DrawFirstHeight(Graphics g)
        {
            Solid(g, 1, 5, 4, 10);
            Light(g, 6, 1, 4, 14);
            Light(g, 11, 10, 4, 5);
            HLine(g, 0, 16, 5);
        }

        private static void DrawFirstWidth(Graphics g)
        {
            Solid(g, 1, 1, 10, 4);
            Light(g, 1, 6, 14, 4);
            Light(g, 1, 11, 5, 4);
            VLine(g, 11, 0, 16);
        }

        // ---- 32x32 menu button --------------------------------------------------------

        private static void DrawMenu(Graphics g)
        {
            Light(g, 3, 13, 10, 15);
            Dashed(g, 3, 4, 10, 24);
            Solid(g, 17, 4, 12, 24);
        }

        // ---- primitives ---------------------------------------------------------------

        private static void Solid(Graphics g, float x, float y, float w, float h)
        {
            using (var b = new SolidBrush(Accent))
                g.FillRectangle(b, x, y, w, h);
            using (var p = new Pen(AccentEdge, 1f))
                g.DrawRectangle(p, x + 0.5f, y + 0.5f, w - 1, h - 1);
        }

        private static void Light(Graphics g, float x, float y, float w, float h)
        {
            using (var b = new SolidBrush(AccentLight))
                g.FillRectangle(b, x, y, w, h);
            using (var p = new Pen(Outline, 1f))
                g.DrawRectangle(p, x + 0.5f, y + 0.5f, w - 1, h - 1);
        }

        private static void Dashed(Graphics g, float x, float y, float w, float h)
        {
            using (var p = new Pen(Guide, 1f) { DashPattern = new[] { 2f, 1f } })
                g.DrawRectangle(p, x + 0.5f, y + 0.5f, w - 1, h - 1);
        }

        private static void HLine(Graphics g, float x1, float x2, float y)
        {
            using (var p = new Pen(Guide, 1f) { DashPattern = new[] { 2f, 1f } })
                g.DrawLine(p, x1, y + 0.5f, x2, y + 0.5f);
        }

        private static void VLine(Graphics g, float x, float y1, float y2)
        {
            using (var p = new Pen(Guide, 1f) { DashPattern = new[] { 2f, 1f } })
                g.DrawLine(p, x + 0.5f, y1, x + 0.5f, y2);
        }
    }
}
