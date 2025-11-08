using PdfSharp.Drawing;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.Versioning;

namespace Excel.Report.PrintDocument
{
    [SupportedOSPlatform("windows")]
    static class GraphicsCompatibleExtensions
    {
        internal static void DrawImage(this Graphics gfx, Image gdimg, double x, double y, double width, double height)
        {
            gfx.DrawImage(
                gdimg,
                new RectangleF((float)x, (float)y, (float)width, (float)height)
            );
        }

        internal static void Restore(this Graphics gfx)
        {
            if (gfx is null) throw new ArgumentNullException(nameof(gfx));
            if (!_states.TryGetValue(gfx, out var stack) || stack.Count == 0) return;
            gfx.Restore(stack.Pop());
        }

        internal static void DrawString(this Graphics gfx, string text, XFont font, XBrush brush, XRect layoutRectangle, XStringFormat format)
        {
            if (gfx is null) throw new ArgumentNullException(nameof(gfx));
            using var gfont = font.ToGdiFont();
            using var gbrush = brush.ToGdiBrush();
            using var gfmt = new StringFormat
            {
                Alignment = format.Alignment switch
                {
                    XStringAlignment.Near => StringAlignment.Near,
                    XStringAlignment.Center => StringAlignment.Center,
                    XStringAlignment.Far => StringAlignment.Far,
                    _ => StringAlignment.Near
                },
                LineAlignment = format.LineAlignment switch
                {
                    XLineAlignment.Near => StringAlignment.Near,
                    XLineAlignment.Center => StringAlignment.Center,
                    XLineAlignment.Far => StringAlignment.Far,
                    _ => StringAlignment.Near
                }
            };
            gfmt.FormatFlags |= StringFormatFlags.NoClip;

            //Adjuctment
            var rect = gfx.ToRectGU(layoutRectangle);
            if (gfmt.LineAlignment != StringAlignment.Center) rect.Height += (gfont.Height / 4);

            gfx.DrawString(text ?? string.Empty, gfont, gbrush, rect, gfmt);
        }

        internal static void TranslateTransform(this Graphics gfx, double dx, double dy)
        {
            if (gfx is null) throw new ArgumentNullException(nameof(gfx));
            if (!_states.TryGetValue(gfx, out var stack))
            {
                stack = new Stack<GraphicsState>();
                _states.Add(gfx, stack);
            }
            stack.Push(gfx.Save());
            gfx.TranslateTransform(dx, dy);
        }

        internal static void DrawRectangle(this Graphics gfx, XBrush brush, double x, double y, double width, double height)
        {
            if (gfx is null) throw new ArgumentNullException(nameof(gfx));
            using var gbrush = brush.ToGdiBrush();
            gfx.FillRectangle(gbrush, (float)x, (float)y, (float)width, (float)height);
        }

        internal static void DrawLine(this Graphics gfx, XPen pen, double x1, double y1, double x2, double y2)
        {
            if (gfx is null) throw new ArgumentNullException(nameof(gfx));
            using var gpen = gfx.ToGdiPen(pen);
            gfx.DrawLine(gpen, (float)x1, (float)y1, (float)x2, (float)y2);
        }

        static readonly ConditionalWeakTable<Graphics, Stack<GraphicsState>> _states = new();

        static RectangleF ToRectGU(this Graphics g, XRect r) =>
            new RectangleF(
                (float)r.X,
                (float)r.Y,
               (float)r.Width,
                (float)r.Height
            );

        static Pen ToGdiPen(this Graphics g, XPen pen)
        {
            static int ToByte(double v) =>
                (int)Math.Round(Math.Clamp(v <= 1.0 ? v * 255.0 : v, 0.0, 255.0));

            var c = pen.Color;
            var gc = Color.FromArgb(ToByte(c.A), ToByte(c.R), ToByte(c.G), ToByte(c.B));

            var p = new Pen(gc, (float)pen.Width);
            p.DashStyle = pen.DashStyle switch
            {
                XDashStyle.Solid => DashStyle.Solid,
                XDashStyle.Dash => DashStyle.Dash,
                XDashStyle.Dot => DashStyle.Dot,
                XDashStyle.DashDot => DashStyle.DashDot,
                XDashStyle.DashDotDot => DashStyle.DashDotDot,
                _ => DashStyle.Solid
            };
            return p;
        }

        static Brush ToGdiBrush(this XBrush brush)
        {
            static int ToByte(double v) =>
                (int)Math.Round(Math.Clamp(v <= 1.0 ? v * 255.0 : v, 0.0, 255.0));

            if (brush is XSolidBrush sb)
            {
                var c = sb.Color;
                var gc = Color.FromArgb(ToByte(c.A), ToByte(c.R), ToByte(c.G), ToByte(c.B));
                return new SolidBrush(gc);
            }

            throw new NotSupportedException("Only XSolidBrush is supported.");
        }

        // XFont -> GDI+ Font
        static string GetFamilyName(XFont font)
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

        static Font ToGdiFont(this XFont font)
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

        //TODO
        internal static Image? TryExtractGdiImage(this XImage xi)
        {
            var fld = xi.GetType().GetField("_stream", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
            if (fld?.GetValue(xi) is not Stream s) return null;
            var p = s.Position; s.Position = 0;
            using var img = Image.FromStream(s, true, true);
            s.Position = p;
            return (Image)img.Clone();
        }
    }
}