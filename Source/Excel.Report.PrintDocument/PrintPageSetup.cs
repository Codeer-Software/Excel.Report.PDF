using ClosedXML.Excel;
using Excel.Report.PDF;

namespace Excel.Report.PrintDocument
{
    public class PrintMargins
    {
        public double Left { get; set; }
        public double Right { get; set; }
        public double Top { get; set; }
        public double Bottom { get; set; }
    }

    public class PrintPageSetup
    {
        public PrintMargins Margins { get; set; } = new PrintMargins();
        public double Width { get; set; }
        public double Height { get; set; }

        public PrintPageSetup FromIXLPageSetup(IXLPageSetup pageSetup)
        {
            (var w, var h) = PaperSizeMap.GetPaperSize(pageSetup.PaperSize);
            return new PrintPageSetup
            {
                Margins = new PrintMargins
                {
                    Left = pageSetup.Margins.Left,
                    Right = pageSetup.Margins.Right,
                    Top = pageSetup.Margins.Top,
                    Bottom = pageSetup.Margins.Bottom,
                },
                Width = w.Point,
                Height = h.Point
            };
        }

        internal PageSetup ToPageSetup()
        {
            return new PageSetup
            {
                Width = Width,
                Height = Height,
                Margins = new Margins
                {
                    Left = Margins.Left,
                    Right = Margins.Right,
                    Top = Margins.Top,
                    Bottom = Margins.Bottom,
                }
            };
        }
    }
}
