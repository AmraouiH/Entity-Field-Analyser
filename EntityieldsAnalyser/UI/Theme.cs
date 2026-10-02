using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;

namespace EntityieldsAnalyser.UI
{
    /// <summary>
    /// Fluent 2 inspired design tokens (aligned with the Dynamics 365 / Power Platform look).
    /// Every color, font and icon used by the tool comes from here.
    /// </summary>
    internal static class Theme
    {
        #region Colors
        public static readonly Color Canvas          = Color.FromArgb(245, 245, 245);
        public static readonly Color Surface         = Color.White;
        public static readonly Color SurfaceHover    = Color.FromArgb(245, 245, 245);
        public static readonly Color SurfacePressed  = Color.FromArgb(224, 224, 224);
        public static readonly Color SurfaceDisabled = Color.FromArgb(250, 250, 250);
        public static readonly Color Stroke          = Color.FromArgb(224, 224, 224);
        public static readonly Color StrokeSubtle    = Color.FromArgb(240, 240, 240);
        public static readonly Color StrokeStrong    = Color.FromArgb(209, 209, 209);

        public static readonly Color TextPrimary     = Color.FromArgb(36, 36, 36);
        public static readonly Color TextSecondary   = Color.FromArgb(97, 97, 97);
        public static readonly Color TextTertiary    = Color.FromArgb(138, 138, 138);
        public static readonly Color TextDisabled    = Color.FromArgb(189, 189, 189);
        public static readonly Color TextOnBrand     = Color.White;

        public static readonly Color Brand           = Color.FromArgb(15, 108, 189);
        public static readonly Color BrandHover      = Color.FromArgb(17, 94, 163);
        public static readonly Color BrandPressed    = Color.FromArgb(12, 59, 94);
        public static readonly Color BrandSubtle     = Color.FromArgb(235, 243, 252);
        public static readonly Color BrandSelected   = Color.FromArgb(207, 228, 250);

        public static readonly Color Success         = Color.FromArgb(16, 124, 16);
        public static readonly Color SuccessSubtle   = Color.FromArgb(241, 250, 241);
        public static readonly Color Warning         = Color.FromArgb(188, 75, 9);
        public static readonly Color WarningSubtle   = Color.FromArgb(255, 249, 245);
        public static readonly Color Danger          = Color.FromArgb(197, 15, 31);
        public static readonly Color DangerSubtle    = Color.FromArgb(253, 243, 244);

        /// <summary>Background of a field row that holds no data at all (0% of use).</summary>
        public static readonly Color UnusedRow       = Color.FromArgb(255, 249, 245);

        /// <summary>Categorical palette for charts (Fluent shared colors, readable on white).</summary>
        public static readonly Color[] ChartPalette =
        {
            Color.FromArgb(15, 108, 189),  // blue
            Color.FromArgb(42, 160, 164),  // teal
            Color.FromArgb(147, 115, 192), // lavender
            Color.FromArgb(227, 0, 140),   // magenta
            Color.FromArgb(202, 80, 16),   // orange
            Color.FromArgb(87, 129, 27),   // forest
            Color.FromArgb(99, 124, 239),  // cornflower
            Color.FromArgb(174, 140, 0),   // gold
            Color.FromArgb(177, 70, 194),  // purple
            Color.FromArgb(0, 91, 112),    // dark teal
            Color.FromArgb(218, 59, 1),    // pumpkin
            Color.FromArgb(78, 37, 124),   // grape
            Color.FromArgb(3, 131, 135),   // seafoam
            Color.FromArgb(142, 86, 46),   // brown
            Color.FromArgb(105, 121, 126), // steel
        };
        #endregion

        #region Fonts
        public const string FontFamily = "Segoe UI";

        public static readonly Font Body          = new Font(FontFamily, 9F, FontStyle.Regular);
        public static readonly Font BodyStrong    = new Font("Segoe UI Semibold", 9F, FontStyle.Regular);
        public static readonly Font Caption       = new Font(FontFamily, 8.25F, FontStyle.Regular);
        public static readonly Font Subtitle      = new Font("Segoe UI Semibold", 10.5F, FontStyle.Regular);
        public static readonly Font Metric        = new Font("Segoe UI Semibold", 17F, FontStyle.Regular);
        public static readonly Font MetricSmall   = new Font("Segoe UI Semibold", 14F, FontStyle.Regular);
        #endregion

        #region Layout helpers
        private static float? _scale;

        /// <summary>DPI factor (1.0 at 96 dpi) used to scale hand painted dimensions.</summary>
        public static float ScaleFactor
        {
            get
            {
                if (_scale == null)
                {
                    using (var g = Graphics.FromHwnd(IntPtr.Zero))
                    {
                        _scale = g.DpiX / 96f;
                    }
                }
                return _scale.Value;
            }
        }

        public static int Scale(int value) => (int)Math.Round(value * ScaleFactor);

        public static GraphicsPath RoundedRect(RectangleF bounds, float radius)
        {
            var path = new GraphicsPath();
            var diameter = Math.Min(radius * 2, Math.Min(bounds.Width, bounds.Height));
            if (diameter <= 0)
            {
                path.AddRectangle(bounds);
                return path;
            }

            path.AddArc(bounds.Left, bounds.Top, diameter, diameter, 180, 90);
            path.AddArc(bounds.Right - diameter, bounds.Top, diameter, diameter, 270, 90);
            path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0, 90);
            path.AddArc(bounds.Left, bounds.Bottom - diameter, diameter, diameter, 90, 90);
            path.CloseFigure();
            return path;
        }

        public static void FillRounded(Graphics g, Color fill, RectangleF bounds, float radius)
        {
            using (var path = RoundedRect(bounds, radius))
            using (var brush = new SolidBrush(fill))
            {
                g.FillPath(brush, path);
            }
        }

        public static void DrawRounded(Graphics g, Color stroke, RectangleF bounds, float radius, float width = 1f)
        {
            using (var path = RoundedRect(bounds, radius))
            using (var pen = new Pen(stroke, width))
            {
                g.DrawPath(pen, path);
            }
        }
        #endregion

        #region Icons
        /// <summary>
        /// Glyphs from the Windows icon fonts (Segoe Fluent Icons on Windows 11, Segoe MDL2 Assets on Windows 10).
        /// They ship with Windows, so the plugin does not need to deploy any extra assembly.
        /// </summary>
        public static class Glyph
        {
            public const char Close        = '';
            public const char List         = '';
            public const char Play         = '';
            public const char Export       = '';
            public const char Help         = '';
            public const char Heart        = '';
            public const char Contact      = '';
            public const char Search       = '';
            public const char Filter       = '';
            public const char CheckMark    = '';
            public const char ChevronDown  = '';
            public const char Info         = '';
            public const char Warning      = '';
            public const char Error        = '';
            public const char Completed    = '';
            public const char Personalize  = '';
            public const char Unlock       = '';
            public const char Package      = '';
            public const char Calculator   = '';
            public const char OpenInNew    = '';
            public const char Table        = '';
            public const char ChevronRight = '';
            public const char Database     = '';
            public const char Chart        = '';
            public const char Clock        = '';
        }

        private static string _iconFontName;
        private static readonly Dictionary<string, Bitmap> IconCache = new Dictionary<string, Bitmap>();

        public static string IconFontName
        {
            get
            {
                if (_iconFontName == null)
                {
                    _iconFontName = "Segoe MDL2 Assets";
                    using (var fonts = new InstalledFontCollection())
                    {
                        foreach (var family in fonts.Families)
                        {
                            if (family.Name == "Segoe Fluent Icons")
                            {
                                _iconFontName = family.Name;
                                break;
                            }
                        }
                    }
                }
                return _iconFontName;
            }
        }

        public static Font IconFont(float pixelSize) => new Font(IconFontName, pixelSize, FontStyle.Regular, GraphicsUnit.Pixel);

        /// <summary>Renders a glyph to a transparent bitmap (cached).</summary>
        public static Bitmap Icon(char glyph, Color color, int size = 16)
        {
            var px = Scale(size);
            var key = $"{(int)glyph}-{color.ToArgb()}-{px}";
            Bitmap bitmap;
            if (IconCache.TryGetValue(key, out bitmap))
                return bitmap;

            bitmap = new Bitmap(px, px);
            using (var g = Graphics.FromImage(bitmap))
            using (var font = IconFont(px))
            using (var brush = new SolidBrush(color))
            using (var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
            {
                g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                g.DrawString(glyph.ToString(), font, brush, new RectangleF(0, 1, px, px), format);
            }

            IconCache[key] = bitmap;
            return bitmap;
        }

        /// <summary>Draws a glyph centered in the given bounds.</summary>
        public static void DrawGlyph(Graphics g, char glyph, Color color, RectangleF bounds, float pixelSize)
        {
            var oldHint = g.TextRenderingHint;
            g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
            using (var font = IconFont(pixelSize))
            using (var brush = new SolidBrush(color))
            using (var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
            {
                g.DrawString(glyph.ToString(), font, brush, bounds, format);
            }
            g.TextRenderingHint = oldHint;
        }
        #endregion

        public static Color Blend(Color a, Color b, float amount)
        {
            return Color.FromArgb(
                (int)(a.R + (b.R - a.R) * amount),
                (int)(a.G + (b.G - a.G) * amount),
                (int)(a.B + (b.B - a.B) * amount));
        }
    }
}
