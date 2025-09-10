using ClosedXML.Excel;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace Excel.Report.PDF
{
    public class ExcelConverter : IDisposable
    {
        public static int MaxRow = 2000;
        public static int MaxColumn = 256;

        public static MemoryStream ConvertToPdf(string filePath)
        {
            using (var converter = new ExcelConverter(filePath))
                return converter.ConvertToPdf();
        }

        public static MemoryStream ConvertToPdf(Stream stream)
        {
            using (var converter = new ExcelConverter(stream))
                return converter.ConvertToPdf();
        }

        public static MemoryStream ConvertToPdf(string filePath, int sheetPosition, PageBreakInfo? pageBreakInfo = null)
        {
            using (var converter = new ExcelConverter(filePath))
                return converter.ConvertToPdf(sheetPosition, pageBreakInfo);
        }

        public static MemoryStream ConvertToPdf(Stream stream, int sheetPosition, PageBreakInfo? pageBreakInfo = null)
        {
            using (var converter = new ExcelConverter(stream))
                return converter.ConvertToPdf(sheetPosition, pageBreakInfo);
        }

        public static MemoryStream ConvertToPdf(string filePath, string sheetName, PageBreakInfo? pageBreakInfo = null)
        {
            using (var converter = new ExcelConverter(filePath))
                return converter.ConvertToPdf(sheetName, pageBreakInfo);
        }

        public static MemoryStream ConvertToPdf(Stream stream, string sheetName, PageBreakInfo? pageBreakInfo = null)
        {
            using (var converter = new ExcelConverter(stream))
                return converter.ConvertToPdf(sheetName, pageBreakInfo);
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

        public MemoryStream ConvertToPdf(int sheetPosition, PageBreakInfo? pageBreakInfo = null)
        {
            using (var pdf = new PdfDocument())
            {
                var ps = _openClosedXML.GetPageSetup(sheetPosition);
                var page = pdf.AddPageEx(ps);
                var allCells = _openClosedXML.GetCellInfo(sheetPosition, page.Width.Point, page.Height.Point, out var scaling, pageBreakInfo);
                return DrawPdf(ps, pdf, page, allCells, scaling);
            }
        }

        public MemoryStream ConvertToPdf(string sheetName, PageBreakInfo? pageBreakInfo = null)
        {
            using (var pdf = new PdfDocument())
            {
                var ps = _openClosedXML.GetPageSetup(sheetName);
                var page = pdf.AddPageEx(ps);
                var allCells = _openClosedXML.GetCellInfo(sheetName, page.Width.Point, page.Height.Point, out var scaling, pageBreakInfo);
                return DrawPdf(ps, pdf, page, allCells, scaling);
            }
        }

        public MemoryStream ConvertToPdf()
        {
            using (var pdf = new PdfDocument())
            {
                foreach (var sheetName in _openClosedXML.GetSheetNames())
                {
                    var ps = _openClosedXML.GetPageSetup(sheetName);
                    var page = pdf.AddPageEx(ps);
                    var allCells = _openClosedXML.GetCellInfo(sheetName, page.Width.Point, page.Height.Point, out var scaling, null);
                    DrawPdfCore(ps, pdf, page, allCells, scaling);
                }
                var outStream = new MemoryStream();
                pdf.Save(outStream);
                return outStream;
            }
        }

        MemoryStream DrawPdf(IXLPageSetup ps, PdfDocument pdf, PdfPage pageSrc, List<List<CellInfo>> allCells, double scaling)
        {
            DrawPdfCore(ps, pdf, pageSrc, allCells, scaling);
            var outStream = new MemoryStream();
            pdf.Save(outStream);
            return outStream;
        }

        void DrawPdfCore(IXLPageSetup ps, PdfDocument pdf, PdfPage pageSrc, List<List<CellInfo>> allCells, double scaling)
        {
            PdfPage? page = pageSrc;
            for (int i = 0; i < allCells.Count; i++)
            {
                if (page == null) page = pdf.AddPageEx(ps);
                using var gfx = XGraphics.FromPdfPage(page);
                page = null;
                var drawLineCache = new DrawLineCache(gfx);

                // Since there are duplicate parts, the loops are separated to prevent overwriting.
                foreach (var cellInfo in allCells[i])
                {
                    FillCellBackColor(gfx, cellInfo);
                }
                foreach (var cellInfo in allCells[i])
                {
                    DrawRuledLine(drawLineCache, scaling, cellInfo);
                }
                foreach (var cellInfo in allCells[i])
                {
                    DrawCellText(gfx, scaling, cellInfo);
                }

                var pictureInfoAndCellInfo = new List<PictureInfoAndCellInfo>();
                foreach (var cellInfo in allCells[i])
                {
                    foreach(var e in cellInfo.Pictures)
                    {
                        pictureInfoAndCellInfo.Add(new PictureInfoAndCellInfo(e, cellInfo));
                    }
                }
                foreach (var e in pictureInfoAndCellInfo.OrderBy(e => e.PictureInfo.Index))
                {
                    DrawPictures(gfx, e);
                }
            }
        }

        void FillCellBackColor(XGraphics gfx, CellInfo cellInfo)
        {
            var cell = cellInfo.Cell!;
            if (cellInfo.MergedFirstCell != null) cell = cellInfo.MergedFirstCell.Cell!;

            var xBackColor = _openClosedXML.ChangeColor(cell.Style.Fill.BackgroundColor);
            if (xBackColor != null)
            {
                var brush = new XSolidBrush(xBackColor.Value);
                gfx.DrawRectangle(brush, cellInfo.X, cellInfo.Y, cellInfo.Width, cellInfo.Height);
            }
        }

        // If you draw two lines in the same place, it will be darker, so skip the second one.
        class DrawLineCache
        {
            XGraphics _gfx;
            Dictionary<string, bool> _cache = new Dictionary<string, bool>();
            public DrawLineCache(XGraphics gfx) => _gfx = gfx;

            public void DrawLine(XPen xPen, double x1, double y1, double x2, double y2)
            {
                var key = $"({Math.Min(x1, x2)},{Math.Min(y1, y2)}),({Math.Max(x1, x2)},{Math.Max(y1, y2)})";
                if (_cache.ContainsKey(key)) return;
                _cache.Add(key, true);
                _gfx.DrawLine(xPen, x1, y1, x2, y2);
            }
        }

        enum Side { Top, Right, Bottom, Left }

        void DrawRuledLine(DrawLineCache gfx, double scaling, CellInfo cellInfo)
        {
            var cell = cellInfo.Cell!;

            // Draw guards for merged ranges
            static bool IsDrawTop(CellInfo i) => i.MergedFirstCell == null || i.Cell?.Address.RowNumber == i.MergedFirstCell.Cell?.Address.RowNumber;
            static bool IsDrawLeft(CellInfo i) => i.MergedFirstCell == null || i.Cell?.Address.ColumnNumber == i.MergedFirstCell.Cell?.Address.ColumnNumber;
            static bool IsDrawBottom(CellInfo i) => i.MergedLastCell == null || i.Cell?.Address.RowNumber == i.MergedLastCell.Cell?.Address.RowNumber;
            static bool IsDrawRight(CellInfo i) => i.MergedLastCell == null || i.Cell?.Address.ColumnNumber == i.MergedLastCell.Cell?.Address.ColumnNumber;

            // ---- Border precedence like Excel (higher wins on shared edges)
            static int Rank(XLBorderStyleValues s) => s switch
            {
                XLBorderStyleValues.Double => 500,
                XLBorderStyleValues.Thick => 400,
                XLBorderStyleValues.Medium => 300,
                XLBorderStyleValues.MediumDashed => 300,
                XLBorderStyleValues.MediumDashDot => 300,
                XLBorderStyleValues.MediumDashDotDot => 300,
                XLBorderStyleValues.Thin => 200,
                XLBorderStyleValues.Dashed => 200,
                XLBorderStyleValues.Dotted => 200,
                XLBorderStyleValues.DashDot => 200,
                XLBorderStyleValues.DashDotDot => 200,
                XLBorderStyleValues.Hair => 100,
                _ => 0
            };

            // Get the neighbor's opposite border on the shared edge
            static XLBorderStyleValues NeighborStyle(IXLCell me, Side side, out XLColor color)
            {
                var ws = me.Worksheet;
                int r = me.Address.RowNumber;
                int c = me.Address.ColumnNumber;
                color = XLColor.Black;

                IXLCell? nb = null;
                XLBorderStyleValues s = XLBorderStyleValues.None;

                switch (side)
                {
                    case Side.Left:
                        if (c > 1) nb = ws.Cell(r, c - 1);
                        if (nb != null) { s = nb.Style.Border.RightBorder; color = nb.Style.Border.RightBorderColor; }
                        break;
                    case Side.Right:
                        nb = ws.Cell(r, c + 1);
                        if (nb != null) { s = nb.Style.Border.LeftBorder; color = nb.Style.Border.LeftBorderColor; }
                        break;
                    case Side.Top:
                        if (r > 1) nb = ws.Cell(r - 1, c);
                        if (nb != null) { s = nb.Style.Border.BottomBorder; color = nb.Style.Border.BottomBorderColor; }
                        break;
                    case Side.Bottom:
                        nb = ws.Cell(r + 1, c);
                        if (nb != null) { s = nb.Style.Border.TopBorder; color = nb.Style.Border.TopBorderColor; }
                        break;
                }
                return s;
            }

            // Decide whether we should draw this shared edge
            static bool ShouldDrawShared(IXLCell me, Side side, XLBorderStyleValues myStyle)
            {
                var nbStyle = NeighborStyle(me, side, out _);
                int myRank = Rank(myStyle);
                int nbRank = Rank(nbStyle);

                if (myRank > nbRank) return true;               // we win -> draw
                if (myRank < nbRank) return false;              // we lose -> skip

                // tie: draw only for Right/Bottom to avoid double painting
                return side == Side.Right || side == Side.Bottom;
            }

            void DrawSide(
                XLBorderStyleValues style, XLColor color, Side side,
                double x1, double y1, double x2, double y2, bool guard)
            {
                if (!guard || style == XLBorderStyleValues.None) return;
                if (!ShouldDrawShared(cell, side, style)) return;

                if (style == XLBorderStyleValues.Double)
                {
                    // Excel-like "Double": two THIN strokes separated by a THIN-sized gap.
                    // Do NOT draw a center line. That would be eaten by a neighbor single line.
                    var thin = _openClosedXML.ConvertToXPen(XLBorderStyleValues.Thin, color, scaling);

                    // Ensure a visible gap on screen/PDF rasterizers
                    double w = Math.Max(thin.Width, 0.7); // >=0.5pt guard for visibility

                    switch (side)
                    {
                        case Side.Top:
                        case Side.Bottom:
                            gfx.DrawLine(thin, x1, y1 - w, x2, y2 - w);
                            gfx.DrawLine(thin, x1, y1 + w, x2, y2 + w);
                            break;
                        case Side.Left:
                        case Side.Right:
                            gfx.DrawLine(thin, x1 - w, y1, x2 - w, y2);
                            gfx.DrawLine(thin, x1 + w, y1, x2 + w, y2);
                            break;
                    }
                    return;
                }

                // Other styles: use the normal pen
                var pen = _openClosedXML.ConvertToXPen(style, color, scaling);
                gfx.DrawLine(pen, x1, y1, x2, y2);
            }

            // Top
            DrawSide(
                cell.Style.Border.TopBorder, cell.Style.Border.TopBorderColor, Side.Top,
                cellInfo.X, cellInfo.Y, cellInfo.X + cellInfo.Width, cellInfo.Y, IsDrawTop(cellInfo));

            // Right
            DrawSide(
                cell.Style.Border.RightBorder, cell.Style.Border.RightBorderColor, Side.Right,
                cellInfo.X + cellInfo.Width, cellInfo.Y, cellInfo.X + cellInfo.Width, cellInfo.Y + cellInfo.Height, IsDrawRight(cellInfo));

            // Bottom
            DrawSide(
                cell.Style.Border.BottomBorder, cell.Style.Border.BottomBorderColor, Side.Bottom,
                cellInfo.X + cellInfo.Width, cellInfo.Y + cellInfo.Height, cellInfo.X, cellInfo.Y + cellInfo.Height, IsDrawBottom(cellInfo));

            // Left
            DrawSide(
                cell.Style.Border.LeftBorder, cell.Style.Border.LeftBorderColor, Side.Left,
                cellInfo.X, cellInfo.Y + cellInfo.Height, cellInfo.X, cellInfo.Y, IsDrawLeft(cellInfo));
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
                    switch (cell.DataType)
                    {
                        case XLDataType.Number:
                        case XLDataType.DateTime:
                            format.Alignment = XStringAlignment.Far;
                            break;
                        case XLDataType.Boolean:
                            format.Alignment = XStringAlignment.Center; 
                            break;
                        default:
                            format.Alignment = XStringAlignment.Near; 
                            break;
                    }
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
            var fontStyle = XFontStyleEx.Regular;
            if (cell.Style.Font.Bold) fontStyle |= XFontStyleEx.Bold;
            if (cell.Style.Font.Italic) fontStyle |= XFontStyleEx.Italic;
            if (cell.Style.Font.Underline != XLFontUnderlineValues.None) fontStyle |= XFontStyleEx.Underline;
            var font = new XFont(cell.Style.Font.FontName, fontSize * scaling, fontStyle);

            var text = cell.GetFormattedString();
            var xFontColor = _openClosedXML.ChangeColor(cell.Style.Font.FontColor) ?? XColor.FromArgb(255, 0, 0, 0);
            var brush = new XSolidBrush(xFontColor);

            double w = cellInfo.MergedWidth != 0 ? cellInfo.MergedWidth : cellInfo.Width;
            double h = cellInfo.MergedHeight != 0 ? cellInfo.MergedHeight : cellInfo.Height;

            // Excel-like padding
            var cellPaddingPt = OpenClosedXML.PixelToPoint(fontSize * (1.0 / 4.0));
            var offset = cellPaddingPt * scaling;
            if (offset * 2 < w) w -= offset * 2;
            if (offset * 2 < h) h -= offset * 2;

            var rect = new XRect(cellInfo.X + offset, cellInfo.Y + offset, w, h);

            var lines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

            // ===== Rotation & vertical text =====
            int raw = cell.Style.Alignment.TextRotation;

            if (raw == 255)
            {
                // Excel's "Vertical Text" (stack)
                DrawVerticalStack(gfx, font, brush, rect, format, lines);
                return;
            }

            // Excel (0..90 = counterclockwise / 91..180 = clockwise (= negative angle))
            // PDFsharp uses positive angles as clockwise, so map as follows
            int anglePdf = 0;
            if (raw <= 90) anglePdf = -raw;        // Up-left slant (Excel +) → negative angle in PDF
            else anglePdf = 180 - raw;    // Up-right slant (Excel -) → positive angle in PDF

            if (anglePdf != 0)
            {
                DrawRotated(gfx, font, brush, rect, format, lines, anglePdf);
                return;
            }

            // ===== Horizontal text (no rotation) =====
            double startY = rect.Y;
            if (format.LineAlignment == XLineAlignment.Center)
                startY += (rect.Height - lines.Length * font.Height) / 2.0;
            else if (format.LineAlignment == XLineAlignment.Far)
                startY += rect.Height - lines.Length * font.Height;

            foreach (var line in lines)
            {
                gfx.DrawString(line, font, brush, new XRect(rect.X, startY, rect.Width, font.Height), format);
                startY += font.Height;
            }

            // ======== Local functions ========

            // Vertical writing (Excel stack): place characters top→bottom, advance columns left→right
            void DrawVerticalStack(XGraphics g, XFont f, XBrush b, XRect r, XStringFormat fmt, string[] cols)
            {
                double step = f.Height;                 // one cell
                double totalW = cols.Length * step;

                double startX = r.X;
                if (fmt.Alignment == XStringAlignment.Center)
                    startX += Math.Max(0, (r.Width - totalW) / 2.0);
                else if (fmt.Alignment == XStringAlignment.Far)
                    startX += Math.Max(0, r.Width - totalW);

                var charFmt = new XStringFormat { Alignment = XStringAlignment.Center, LineAlignment = XLineAlignment.Near };

                for (int c = 0; c < cols.Length; c++)
                {
                    string col = cols[c] ?? string.Empty;
                    double colH = col.Length * step;

                    double y = r.Y;
                    if (fmt.LineAlignment == XLineAlignment.Center)
                        y += Math.Max(0, (r.Height - colH) / 2.0);
                    else if (fmt.LineAlignment == XLineAlignment.Far)
                        y += Math.Max(0, r.Height - colH);

                    double x = startX + c * step;

                    for (int i = 0; i < col.Length; i++)
                    {
                        string ch = col.Substring(i, 1);
                        g.DrawString(ch, f, b, new XRect(x, y + i * step, step, step), charFmt);
                    }
                }
            }

            // Arbitrary-angle drawing: rotate the coordinate system around the rectangle center (do not swap width/height)
            void DrawRotated(XGraphics g, XFont f, XBrush b, XRect r, XStringFormat fmt, string[] content, int angle)
            {
                g.Save();

                // Rotate about the center (PDFsharp uses positive = clockwise)
                g.TranslateTransform(r.X + r.Width / 2.0, r.Y + r.Height / 2.0);
                g.RotateTransform(angle);

                var rr = new XRect(-r.Width / 2.0, -r.Height / 2.0, r.Width, r.Height);

                double y = rr.Y;
                if (fmt.LineAlignment == XLineAlignment.Center)
                    y += (rr.Height - content.Length * f.Height) / 2.0;
                else if (fmt.LineAlignment == XLineAlignment.Far)
                    y += rr.Height - content.Length * f.Height;

                foreach (var line in content)
                {
                    g.DrawString(line, f, b, new XRect(rr.X, y, rr.Width, f.Height), fmt);
                    y += f.Height;
                }

                g.Restore();
            }
        }

        class PictureInfoAndCellInfo
        {
            public PictureInfo PictureInfo { get; set; }
            public CellInfo CellInfo { get; set; }
            public PictureInfoAndCellInfo(PictureInfo pictureInfo, CellInfo cellInfo)
            {
                PictureInfo = pictureInfo;
                CellInfo = cellInfo;
            }
        }

        static void DrawPictures(XGraphics gfx, PictureInfoAndCellInfo item)
        {
            var xImage = XImage.FromStream(item.PictureInfo.Picture!);
            gfx.DrawImage(xImage,
            item.CellInfo.X + item.PictureInfo.X,
                item.CellInfo.Y + item.PictureInfo.Y,
                item.PictureInfo.Width,
                item.PictureInfo.Height);
        }
    }
}
