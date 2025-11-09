using ClosedXML.Excel;
using Excel.Report.PDF;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.Versioning;

namespace Excel.Report.PrintDocument
{
    [SupportedOSPlatform("windows")]
    class VirtualGraphics : IVirtualGraphics
    {
        List<Action<Graphics>> _actions;
        List<IDisposable> _disposables;
        Stack<GraphicsState> _states = new();

        internal VirtualGraphics(List<Action<Graphics>> actions, List<IDisposable> disposables)
        {
            _actions = actions;
            _disposables = disposables;
        }

        public void DrawImage(MemoryStream stream, double x, double y, double width, double height)
        {
            var image = System.Drawing.Image.FromStream(stream);
            _disposables.Add(image);
            _actions.Add(g => g.DrawImage(image, x, y, width, height));
        }
        public void Restore()
            => _actions.Add(g => g.Restore(_states));

        public void DrawString(string text, VirtualFont font, VirtualColor brush, VirtualRect layoutRectangle, VirtualStringFormat format)
            => _actions.Add(g => g.DrawString(text, font, brush, layoutRectangle, format));

        public void TranslateTransform(double dx, double dy)
            => _actions.Add(g => g.TranslateTransform(_states, dx, dy));

        public void DrawRectangle(VirtualColor brush, double x, double y, double width, double height)
            => _actions.Add(g => g.DrawRectangle(brush, x, y, width, height));

        public void DrawLine(VirtualPen pen, double x1, double y1, double x2, double y2)
            => _actions.Add(g => g.DrawLine(pen, x1, y1, x2, y2));

        public void Save()
            => _actions.Add(g => g.Save());

        public void RotateTransform(int angle)
            => _actions.Add(g => g.RotateTransform(angle));

        public double GetFontHeight(VirtualFont vf)
        {
            var pt = GetLineHeightInPoints(vf.ToGdiFont());
            return pt;
        }

        float GetLineHeightInPoints(Font f)
        {
            var fam = f.FontFamily;
            var style = f.Style;
            int em = fam.GetEmHeight(style);       // デザイン単位の em 高さ
            int line = fam.GetLineSpacing(style);    // デザイン単位の行送り
            return f.SizeInPoints * line / em;        // ポイント値（DPI不要）
        }
    }

    [SupportedOSPlatform("windows")]
    class VirtualPage : IVirtualPage
    {
        List<IDisposable> _disposables = new();
        List<Action<Graphics>> _actions = new();
        public IXLPageSetup PageSetup { get; }
        public VirtualPage(IXLPageSetup ps) => PageSetup = ps;
        public IVirtualGraphics CreateGraphics() => new VirtualGraphics(_actions, _disposables);
        public void DrawTo(Graphics g)
        {
            _actions.ForEach(a => a(g));
            _disposables.ForEach(d => d.Dispose());
        }
    }

    [SupportedOSPlatform("windows")]
    class VirtualDocument : IVirtualDocument
    {
        readonly List<VirtualPage> _pages = new();
        public int PageCount => _pages.Count;
        public IVirtualPage AddPage(IXLPageSetup ps)
        {
            var page = new VirtualPage(ps);
            _pages.Add(page);
            return page;
        }

        public void DrawTo(Graphics g, int page)=> _pages[page].DrawTo(g);
    }
}
