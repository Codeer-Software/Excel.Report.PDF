using ClosedXML.Excel;
using Excel.Report.PDF;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.Versioning;

namespace Excel.Report.PrintDocument
{
    [SupportedOSPlatform("windows")]
    static class GraphicsCompatibleExtensions
    {
        internal static void TranslateTransform(this Graphics gfx, Stack<GraphicsState> states, double dx, double dy)
        {
            if (gfx is null) throw new ArgumentNullException(nameof(gfx));
            states.Push(gfx.Save());
            gfx.TranslateTransform((float)dx, (float)dy);
        }

        internal static void Restore(this Graphics gfx, Stack<GraphicsState> states)
        {
            if (gfx is null) throw new ArgumentNullException(nameof(gfx));
            gfx.Restore(states.Pop());
        }

        internal static void DrawImage(this Graphics gfx, Image gdimg, double x, double y, double width, double height)
            => gfx.DrawImage(gdimg, new RectangleF((float)x, (float)y, (float)width, (float)height));

        internal static void DrawString(this Graphics gfx, string text, VirtualFont font, VirtualColor brush, VirtualRect layoutRectangle, VirtualStringFormat format)
        {
            if (gfx is null) throw new ArgumentNullException(nameof(gfx));
            var rect = new RectangleF((float)layoutRectangle.X, (float)layoutRectangle.Y, (float)layoutRectangle.Width, (float)layoutRectangle.Height);
            using var gfont = font.ToGdiFont();
            using var gbrush = brush.ToGdiBrush();
            using var gfmt = new StringFormat
            {
                Alignment = format.Alignment switch
                {
                    VirtualAlignment.Near => StringAlignment.Near,
                    VirtualAlignment.Center => StringAlignment.Center,
                    VirtualAlignment.Far => StringAlignment.Far,
                    _ => StringAlignment.Near
                },
                LineAlignment = format.LineAlignment switch
                {
                    VirtualAlignment.Near => StringAlignment.Near,
                    VirtualAlignment.Center => StringAlignment.Center,
                    VirtualAlignment.Far => StringAlignment.Far,
                    _ => StringAlignment.Near
                }
            };
            gfmt.FormatFlags |= StringFormatFlags.NoClip;

            //Adjuctment
            if (gfmt.LineAlignment != StringAlignment.Center) rect.Height += (gfont.Height / 4);

            gfx.DrawString(text ?? string.Empty, gfont, gbrush, rect, gfmt);
        }

        internal static void DrawRectangle(this Graphics gfx, VirtualColor brush, double x, double y, double width, double height)
        {
            if (gfx is null) throw new ArgumentNullException(nameof(gfx));
            using var gbrush = brush.ToGdiBrush();
            gfx.FillRectangle(gbrush, (float)x, (float)y, (float)width, (float)height);
        }

        internal static void DrawLine(this Graphics gfx, VirtualPen pen, double x1, double y1, double x2, double y2)
        {
            if (gfx is null) throw new ArgumentNullException(nameof(gfx));
            using var gpen = gfx.ToGdiPen(pen);
            gfx.DrawLine(gpen, (float)x1, (float)y1, (float)x2, (float)y2);
        }

        static Pen ToGdiPen(this Graphics g, VirtualPen pen)
        {
            static int ToByte(double v) =>
                (int)Math.Round(Math.Clamp(v <= 1.0 ? v * 255.0 : v, 0.0, 255.0));

            var c = pen.Color;
            var gc = Color.FromArgb(ToByte(c.A), ToByte(c.R), ToByte(c.G), ToByte(c.B));

            var p = new Pen(gc, (float)pen.Width);
            switch (pen.BorderStyleValues)
            {
                case XLBorderStyleValues.Dotted:
                    p.DashStyle = DashStyle.Dot;
                    break;
                case XLBorderStyleValues.Dashed:
                    p.DashStyle = DashStyle.Dash;
                    break;
                case XLBorderStyleValues.Hair:
                    // PDFsharp doesn't have direct support for the Hair style, so we approximate it as a Dot.
                    p.DashStyle = DashStyle.Dot;
                    break;
                case XLBorderStyleValues.MediumDashed:
                    p.DashStyle = DashStyle.Dash;
                    break;
                case XLBorderStyleValues.DashDot:
                    p.DashStyle = DashStyle.DashDot;
                    break;
                case XLBorderStyleValues.MediumDashDot:
                case XLBorderStyleValues.SlantDashDot:
                case XLBorderStyleValues.MediumDashDotDot:
                    p.DashStyle = DashStyle.DashDotDot;
                    break;
                case XLBorderStyleValues.None:
                case XLBorderStyleValues.Thin:
                case XLBorderStyleValues.Medium:
                case XLBorderStyleValues.Thick:
                case XLBorderStyleValues.Double:
                default:
                    p.DashStyle = DashStyle.Solid;
                    break;
            }
            return p;
        }

        static Brush ToGdiBrush(this VirtualColor c)
        {
            static int ToByte(double v) => (int)Math.Round(Math.Clamp(v <= 1.0 ? v * 255.0 : v, 0.0, 255.0));
            var gc = Color.FromArgb(ToByte(c.A), ToByte(c.R), ToByte(c.G), ToByte(c.B));
            return new SolidBrush(gc);
        }

        internal static double GetFontHeight(this Graphics gfx, VirtualFont font)
        {
            using var gfont = font.ToGdiFont();
            return gfont.GetHeight(gfx);
        }

        static Font ToGdiFont(this VirtualFont font)
        {
            var style = FontStyle.Regular;

            var t = font.GetType();
            bool Has(string name) => t.GetProperty(name)?.GetValue(font) as bool? ?? false;

            if (Has("Bold")) style |= FontStyle.Bold;
            if (Has("Italic")) style |= FontStyle.Italic;
            if (Has("Underline")) style |= FontStyle.Underline;
            if (Has("Strikeout")) style |= FontStyle.Strikeout;

            var family = GetFamilyName(font);

            var sizeProp = t.GetProperty("Size");
            var size = (double)(sizeProp?.GetValue(font) ?? 12.0);

            // Font size is pt
            return new Font(family, (float)size, style, GraphicsUnit.Point);
        }

        static string GetFamilyName(VirtualFont font)
        {
            var t = font.GetType();
            var p = t.GetProperty("Name");
            if (p?.GetValue(font) is string s && !string.IsNullOrEmpty(s)) return s;

            var ff = t.GetProperty("FontFamily")?.GetValue(font);
            var ffName = ff?.GetType().GetProperty("Name")?.GetValue(ff) as string;
            if (!string.IsNullOrEmpty(ffName)) return ffName;

            var fam = t.GetProperty("FamilyName")?.GetValue(font) as string;
            if (!string.IsNullOrEmpty(fam)) return fam;

            return "Segoe UI";
        }
    }
}