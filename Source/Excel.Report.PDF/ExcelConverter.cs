using ClosedXML.Excel;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace Excel.Report.PDF
{
    public class ExcelConverter : IDisposable
    {
        public static MemoryStream ConvertToPdf(string filePath, int sheetPosition)
        {
            using (var converter = new ExcelConverter(filePath))
                return converter.ConvertToPdf(sheetPosition);
        }

        public static MemoryStream ConvertToPdf(Stream stream, int sheetPosition)
        {
            using (var converter = new ExcelConverter(stream))
                return converter.ConvertToPdf(sheetPosition);
        }

        public static MemoryStream ConvertToPdf(string filePath, string sheetName)
        {
            using (var converter = new ExcelConverter(filePath))
                return converter.ConvertToPdf(sheetName);
        }

        public static MemoryStream ConvertToPdf(Stream stream, string sheetName)
        {
            using (var converter = new ExcelConverter(stream))
                return converter.ConvertToPdf(sheetName);
        }

        OpenClosedXML _openClosedXML;
        Stream? _myOpenStream;

        public ExcelConverter(Stream stream)
            => _openClosedXML = new OpenClosedXML(stream);

        public ExcelConverter(string path)
        {
            _myOpenStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            _openClosedXML = new OpenClosedXML(_myOpenStream);
        }

        public void Dispose()
        {
            _openClosedXML.Dispose();
            _myOpenStream?.Dispose();
        }

        public MemoryStream ConvertToPdf(int sheetPosition)
        {
            using (var pdf = new PdfDocument())
            {
                var page = pdf.AddPage();
                if (_openClosedXML.IsLandscape(sheetPosition)) page.Orientation = PdfSharp.PageOrientation.Landscape;
                var allCells = _openClosedXML.GetCellInfo(sheetPosition, page.Width.Point, page.Height.Point, out var scaling);
                return DrawPdf(pdf, page, allCells, scaling);
            }
        }

        public MemoryStream ConvertToPdf(string sheetName)
        {
            using (var pdf = new PdfDocument())
            {
                var page = pdf.AddPage();
                if (_openClosedXML.IsLandscape(sheetName)) page.Orientation = PdfSharp.PageOrientation.Landscape;
                var allCells = _openClosedXML.GetCellInfo(sheetName, page.Width.Point, page.Height.Point, out var scaling);
                return DrawPdf(pdf, page, allCells, scaling);
            }
        }

        MemoryStream DrawPdf(PdfDocument pdf, PdfPage pageSrc, List<List<CellInfo>> allCells, double scaling)
        {
            PdfPage? page = pageSrc;
            for (int i = 0; i < allCells.Count; i++)
            {
                if (page == null) page = pdf.AddPage();
                using var gfx = XGraphics.FromPdfPage(page);
                page = null;

                // Since there are duplicate parts, the loops are separated to prevent overwriting.
                foreach (var cellInfo in allCells[i])
                {
                    FillCellBackColor(gfx, cellInfo);
                }
                foreach (var cellInfo in allCells[i])
                {
                    DrawRuledLine(gfx, scaling, cellInfo);
                }
                foreach (var cellInfo in allCells[i])
                {
                    DrawCellText(gfx, scaling, cellInfo);
                }
                foreach (var cellInfo in allCells[i])
                {
                    DrawPictures(gfx, cellInfo);
                }
            }

            var outStream = new MemoryStream();
            pdf.Save(outStream);
            return outStream;
        }

        void FillCellBackColor(XGraphics gfx, CellInfo cellInfo)
        {
            var cell = cellInfo.Cell!;
            if (cellInfo.MeargedTopCell != null) cell = cellInfo.MeargedTopCell.Cell!;

            var xBackColor = _openClosedXML.ChangeColor(cell.Style.Fill.BackgroundColor);
            if (xBackColor != null)
            {
                var brush = new XSolidBrush(xBackColor.Value);
                gfx.DrawRectangle(brush, cellInfo.X, cellInfo.Y, cellInfo.Width, cellInfo.Height);
            }
        }

        void DrawRuledLine(XGraphics gfx, double scaling, CellInfo cellInfo)
        {
            var cell = cellInfo.Cell!;

            if (cell.Style.Border.TopBorder != XLBorderStyleValues.None)
            {
                gfx.DrawLine(_openClosedXML.ConvertToXPen(cell.Style.Border.TopBorder, cell.Style.Border.TopBorderColor, scaling), cellInfo.X, cellInfo.Y, cellInfo.X + cellInfo.Width, cellInfo.Y);
            }
            if (cell.Style.Border.RightBorder != XLBorderStyleValues.None)
            {
                gfx.DrawLine(_openClosedXML.ConvertToXPen(cell.Style.Border.RightBorder, cell.Style.Border.RightBorderColor, scaling), cellInfo.X + cellInfo.Width, cellInfo.Y, cellInfo.X + cellInfo.Width, cellInfo.Y + cellInfo.Height);
            }
            if (cell.Style.Border.BottomBorder != XLBorderStyleValues.None)
            {
                gfx.DrawLine(_openClosedXML.ConvertToXPen(cell.Style.Border.BottomBorder, cell.Style.Border.BottomBorderColor, scaling), cellInfo.X + cellInfo.Width, cellInfo.Y + cellInfo.Height, cellInfo.X, cellInfo.Y + cellInfo.Height);
            }
            if (cell.Style.Border.LeftBorder != XLBorderStyleValues.None)
            {
                gfx.DrawLine(_openClosedXML.ConvertToXPen(cell.Style.Border.LeftBorder, cell.Style.Border.LeftBorderColor, scaling), cellInfo.X, cellInfo.Y + cellInfo.Height, cellInfo.X, cellInfo.Y);
            }
        }

        void DrawCellText(XGraphics gfx, double scaling, CellInfo cellInfo)
        {
            var cell = cellInfo.Cell!;

            // Alignment
            var format = new XStringFormat();
            switch (cell.Style.Alignment.Horizontal)
            {
                case XLAlignmentHorizontalValues.Center:
                    format.Alignment = XStringAlignment.Center;
                    break;
                case XLAlignmentHorizontalValues.Right:
                    format.Alignment = XStringAlignment.Far;
                    break;
                default:
                    format.Alignment = XStringAlignment.Near;
                    break;
            }
            switch (cell.Style.Alignment.Vertical)
            {
                case XLAlignmentVerticalValues.Center:
                    format.LineAlignment = XLineAlignment.Center;
                    break;
                case XLAlignmentVerticalValues.Bottom:
                    format.LineAlignment = XLineAlignment.Far;
                    break;
                default:
                    format.LineAlignment = XLineAlignment.Near;
                    break;
            }

            // Font
            double fontSize = cell.Style.Font.FontSize;
            XFont font = new XFont(cell.Style.Font.FontName, fontSize * scaling);

            var text = cell.GetFormattedString();
            var xFontColor = _openClosedXML.ChangeColor(cell.Style.Font.FontColor) ?? XColor.FromArgb(255, 0, 0, 0);

            var w = cellInfo.MergedWidth != 0 ? cellInfo.MergedWidth : cellInfo.Width;
            var h = cellInfo.MergedHeight != 0 ? cellInfo.MergedHeight : cellInfo.Height;

            var CellPadding = OpenClosedXML.PixelToPoint(fontSize * (1.0/4.0));
            var offset = CellPadding * scaling;
            if (offset * 2 < w) w = w - offset * 2;
            if (offset * 2 < h) h = h - offset * 2;
            gfx.DrawString(text, font, new XSolidBrush(xFontColor), new XRect(cellInfo.X + offset, cellInfo.Y + offset, w, h), format);
        }

        static void DrawPictures(XGraphics gfx, CellInfo cellInfo)
        {
            foreach (var pictureInfo in cellInfo.Pictures)
            {
                var xImage = XImage.FromStream(pictureInfo.Picture!);
                gfx.DrawImage(xImage,
                cellInfo.X + pictureInfo.X,
                    cellInfo.Y + pictureInfo.X,
                    pictureInfo.Width,
                    pictureInfo.Height);
            }
        }
    }
}
