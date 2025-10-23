using ClosedXML.Excel;
using PdfSharp.Drawing;
using System.Drawing;
using System.Runtime.Versioning;

namespace Excel.Report.PrintDocument
{
    [SupportedOSPlatform("windows")]
    class VirtualGraphics
    {
        List<Action<Graphics>> _actions;

        internal VirtualGraphics(List<Action<Graphics>> actions)
            => _actions = actions;

        internal void DrawImage(XImage image, double x, double y, double width, double height)
            => _actions.Add(g => g.DrawImage(image, x, y, width, height));

        internal void Restore()
            => _actions.Add(g => g.ResetTransform());

        internal void DrawString(string text, XFont font, XBrush brush, XRect layoutRectangle, XStringFormat format)
            => _actions.Add(g => g.DrawString(text, font, brush, layoutRectangle, format));

        internal void TranslateTransform(double dx, double dy)
            => _actions.Add(g => g.TranslateTransform(dx, dy));

        internal void DrawRectangle(XBrush brush, double x, double y, double width, double height)
            => _actions.Add(g => g.DrawRectangle(brush, x, y, width, height));

        internal void DrawLine(XPen pen, double x1, double y1, double x2, double y2)
            => _actions.Add(g => g.DrawLine(pen, x1, y1, x2, y2));

        internal void Save()
            => _actions.Add(g => g.Save());

        internal void RotateTransform(int angle)
            => _actions.Add(g => g.RotateTransform(angle));
    }

    [SupportedOSPlatform("windows")]
    class VirtualPage
    {
        List<Action<Graphics>> _actions = new();
        public XLPaperSize PaperSize { get; }
        public VirtualPage(XLPaperSize size) => PaperSize = size;
        public VirtualGraphics CreateGraphics() => new(_actions);
        public void DrawTo(Graphics g) => _actions.ForEach(a => a(g));
    }

    [SupportedOSPlatform("windows")]
    class VirtualDocument
    {
        readonly List<VirtualPage> _pages = new();
        public int PageCount => _pages.Count;
        public VirtualPage AddPage(IXLPageSetup ps)
        {
            var page = new VirtualPage(ps.PaperSize);
            _pages.Add(page);
            return page;
        }

        public void DrawTo(Graphics g, int page)=> _pages[page].DrawTo(g);
    }
}
