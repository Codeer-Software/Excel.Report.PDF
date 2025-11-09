using ClosedXML.Excel;
using Excel.Report.PDF;
using PdfSharp.Drawing;

namespace Excel.Report.PrintDocument
{
    class PdfVirtualGraphics : IVirtualGraphics
    {
        List<Action<XGraphics>> _actions;
        List<IDisposable> _disposables;

        internal PdfVirtualGraphics(List<Action<XGraphics>> actions, List<IDisposable> disposables)
        {
            _actions = actions;
            _disposables = disposables;
        }

        public void DrawImage(MemoryStream stream, double x, double y, double width, double height)
        {
            var image = XImage.FromStream(stream);
            _disposables.Add(image);
            _actions.Add(g => g.DrawImage(image, x, y, width, height));
        }

        public void Restore() => _actions.Add(g => g.Restore());

        public void DrawString(string text, VirtualFont vFont, VirtualColor vBrush, VirtualRect vRect, VirtualStringFormat vFormat)
        {
            var font = ConvertToXFont(vFont);
            var brush = new XSolidBrush(ConvertToXColor(vBrush));
            var rect = new XRect(vRect.X, vRect.Y, vRect.Width, vRect.Height);
            var format = ConvertToXStringFormat(vFormat);
            _actions.Add(g => g.DrawString(text, font, brush, rect, format));
        }

        public void TranslateTransform(double dx, double dy)
            => _actions.Add(g => g.TranslateTransform(dx, dy));

        public void DrawRectangle(VirtualColor vBrush, double x, double y, double width, double height)
            => _actions.Add(g => g.DrawRectangle(new XSolidBrush(ConvertToXColor(vBrush)), x, y, width, height));

        public void DrawLine(VirtualPen pen, double x1, double y1, double x2, double y2)
            => _actions.Add(g => g.DrawLine(ConvertToXPen(pen), x1, y1, x2, y2));

        public void Save()
            => _actions.Add(g => g.Save());

        public void RotateTransform(int angle)
            => _actions.Add(g => g.RotateTransform(angle));

        public double GetFontHeight(VirtualFont vFont)
            => ConvertToXFont(vFont).GetHeight();

        static XStringFormat ConvertToXStringFormat(VirtualStringFormat vFormat)
        {
            var format = new XStringFormat();
            format.Alignment = vFormat.Alignment switch
            {
                VirtualAlignment.Near => XStringAlignment.Near,
                VirtualAlignment.Center => XStringAlignment.Center,
                VirtualAlignment.Far => XStringAlignment.Far,
                _ => XStringAlignment.Near
            };
            format.LineAlignment = vFormat.LineAlignment switch
            {
                VirtualAlignment.Near => XLineAlignment.Near,
                VirtualAlignment.Center => XLineAlignment.Center,
                VirtualAlignment.Far => XLineAlignment.Far,
                _ => XLineAlignment.Near
            };
            return format;
        }

        static XFont ConvertToXFont(VirtualFont vFont)
        {
            var fontStyle = XFontStyleEx.Regular;
            if (vFont.Core.Bold) fontStyle |= XFontStyleEx.Bold;
            if (vFont.Core.Italic) fontStyle |= XFontStyleEx.Italic;
            if (vFont.Core.Underline != XLFontUnderlineValues.None) fontStyle |= XFontStyleEx.Underline;

            var font = new XFont(vFont.Core.FontName, vFont.Core.FontSize * vFont.Scaling, fontStyle);
            return font;
        }

        static XColor ConvertToXColor(VirtualColor vColor)
            => XColor.FromArgb(vColor.A, vColor.R, vColor.G, vColor.B);

        static XPen ConvertToXPen(VirtualPen vPen)
        {
            var pen = new XPen(ConvertToXColor(vPen.Color), vPen.Width);
            switch (vPen.BorderStyleValues)
            {
                case XLBorderStyleValues.None:
                    pen.DashStyle = XDashStyle.Solid;
                    pen.Color = XColors.Transparent;
                    break;
                case XLBorderStyleValues.Thin:
                case XLBorderStyleValues.Medium:
                case XLBorderStyleValues.Thick:
                    pen.DashStyle = XDashStyle.Solid;
                    break;
                case XLBorderStyleValues.Dotted:
                    pen.DashStyle = XDashStyle.Dot;
                    break;
                case XLBorderStyleValues.Dashed:
                    pen.DashStyle = XDashStyle.Dash;
                    break;
                case XLBorderStyleValues.Double:
                    // PDFsharp doesn't have direct support for Double style, so we approximate it as Solid.
                    pen.DashStyle = XDashStyle.Solid;
                    break;
                case XLBorderStyleValues.Hair:
                    // PDFsharp doesn't have direct support for the Hair style, so we approximate it as a Dot.
                    pen.DashStyle = XDashStyle.Dot;
                    break;
                case XLBorderStyleValues.MediumDashed:
                    pen.DashStyle = XDashStyle.Dash;
                    break;
                case XLBorderStyleValues.DashDot:
                    pen.DashStyle = XDashStyle.DashDot;
                    break;
                case XLBorderStyleValues.MediumDashDot:
                case XLBorderStyleValues.SlantDashDot:
                case XLBorderStyleValues.MediumDashDotDot:
                    pen.DashStyle = XDashStyle.DashDotDot;
                    break;
                default:
                    pen.DashStyle = XDashStyle.Solid;
                    break;
            }
            return pen;
        }
    }

    class PdfVirtualPage : IVirtualPage
    {
        List<IDisposable> _disposables = new();
        List<Action<XGraphics>> _actions = new();
        public IXLPageSetup PageSetup { get; }
        public PdfVirtualPage(IXLPageSetup pageSetup) => PageSetup = pageSetup;
        public IVirtualGraphics CreateGraphics() => new PdfVirtualGraphics(_actions, _disposables);
        public void DrawTo(XGraphics g)
        {
            _actions.ForEach(a => a(g));
            _disposables.ForEach(d => d.Dispose());
        }
    }

    class PdfVirtualDocument : IVirtualDocument
    {
        public List<PdfVirtualPage> Pages { get; } = new();
        public int PageCount => Pages.Count;
        public IVirtualPage AddPage(IXLPageSetup ps)
        {
            var page = new PdfVirtualPage(ps);
            Pages.Add(page);
            return page;
        }

        public void DrawTo(XGraphics g, int page)=> Pages[page].DrawTo(g);
    }
}
