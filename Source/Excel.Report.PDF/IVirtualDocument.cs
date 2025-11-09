using ClosedXML.Excel;

namespace Excel.Report.PDF
{
    class VirtualColor
    {
        internal int A { get; set; }
        internal int R { get; set; }
        internal int G { get; set; }
        internal int B { get; set; }
        internal VirtualColor(int a, int r, int g, int b)
        {
            A = a;
            R = r;
            G = g;
            B = b;
        }
    }

    class VirtualPen
    {
        internal VirtualColor Color { get; set; }
        internal double Width { get; set; }
        internal XLBorderStyleValues BorderStyleValues { get; set; }
        internal VirtualPen(VirtualColor color, double width, XLBorderStyleValues border)
        {
            Color = color;
            Width = width;
            BorderStyleValues = border;
        }
    }
    enum VirtualAlignment
    {
        Near,
        Center,
        Far
    }

    class VirtualStringFormat
    {
        internal VirtualAlignment Alignment { get; set; }
        internal VirtualAlignment LineAlignment { get; set; }
    }

    class VirtualRect
    {
        internal double X { get; set; }
        internal double Y { get; set; }
        internal double Width { get; set; }
        internal double Height { get; set; }
        internal VirtualRect(double x, double y, double width, double height)
        {
            X = x;
            Y = y;
            Width = width;
            Height = height;
        }
    }

    class VirtualFont
    {
        internal double Scaling { get; set; }
        internal IXLFont Core { get; set; }
        internal VirtualFont(IXLFont core, double scaling)
        {
            Core = core;
            Scaling = scaling;
        }
    }   

    interface IVirtualGraphics
    {
        void DrawImage(MemoryStream stream, double x, double y, double width, double height);
        void Restore();
        void DrawString(string text, VirtualFont font, VirtualColor brush, VirtualRect layoutRectangle, VirtualStringFormat format);
        void TranslateTransform(double dx, double dy);
        void DrawRectangle(VirtualColor brush, double x, double y, double width, double height);
        void DrawLine(VirtualPen pen, double x1, double y1, double x2, double y2);
        void Save();
        void RotateTransform(int angle);
        double GetFontHeight(VirtualFont font);
    }

    interface IVirtualPage
    {
        IXLPageSetup PageSetup { get; }
        IVirtualGraphics CreateGraphics();
    }

    interface IVirtualDocument
    {
        int PageCount { get; }
        IVirtualPage AddPage(IXLPageSetup ps);
    }
}
