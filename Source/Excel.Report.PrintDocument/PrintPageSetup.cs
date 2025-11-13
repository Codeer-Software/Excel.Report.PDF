using ClosedXML.Excel;
using Excel.Report.PDF;

namespace Excel.Report.PrintDocument
{
    public class PrintMargins
    {
        public double LeftPoint { get; set; }
        public double RightPoint { get; set; }
        public double TopPoint { get; set; }
        public double BottomPoint { get; set; }
    }

    public class PrintPageSetup
    {
        public PrintMargins Margins { get; set; } = new PrintMargins();
        public double WidthPoint { get; set; }
        public double HeightPoint { get; set; }

        public PrintPageSetup FromIXLPageSetup(IXLPageSetup pageSetup)
        {
            (var w, var h) = PaperSizeMap.GetPaperSize(pageSetup.PaperSize);
            return new PrintPageSetup
            {
                Margins = new PrintMargins
                {
                    LeftPoint = pageSetup.Margins.Left,
                    RightPoint = pageSetup.Margins.Right,
                    TopPoint = pageSetup.Margins.Top,
                    BottomPoint = pageSetup.Margins.Bottom,
                },
                WidthPoint = w.Point,
                HeightPoint = h.Point
            };
        }

        public static double MmToPoint(double mm)
            => mm * 72.0 / 25.4;

        internal PageSetup ToPageSetup()
        {
            return new PageSetup
            {
                Width = WidthPoint,
                Height = HeightPoint,
                Margins = new Margins
                {
                    Left = Margins.LeftPoint,
                    Right = Margins.RightPoint,
                    Top = Margins.TopPoint,
                    Bottom = Margins.BottomPoint,
                }
            };
        }
    }
}
