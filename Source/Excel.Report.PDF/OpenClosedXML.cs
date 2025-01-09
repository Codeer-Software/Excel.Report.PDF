using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;
using PdfSharp.Drawing;
using ClosedXML.Excel.Drawings;
using DocumentFormat.OpenXml.Spreadsheet;
using Color = System.Drawing.Color;

namespace Excel.Report.PDF
{
    class OpenClosedXML : IDisposable
    {
        readonly SpreadsheetDocument _document;

        internal XLWorkbook Workbook { get; }

        internal bool IsLandscape(int sheetPosition)
            => Workbook.Worksheet(sheetPosition).PageSetup.PageOrientation == XLPageOrientation.Landscape;

        internal bool IsLandscape(string sheetName)
            => Workbook.Worksheet(sheetName).PageSetup.PageOrientation == XLPageOrientation.Landscape;

        internal OpenClosedXML(Stream stream)
        {
            stream.Position = 0;
            _document = SpreadsheetDocument.Open(stream, false);
            stream.Position = 0;
            Workbook = new XLWorkbook(stream);
        }

        public void Dispose()
        {
            _document.Dispose();
            Workbook.Dispose();
        }

        WorksheetPart GetWorkSheetPartByPosition(int sheetPosition)
        {
            var workbookPart = _document.WorkbookPart;
            if (workbookPart == null) throw new InvalidDataException("Invalid sheet"); 
            var sheet = workbookPart.Workbook.Sheets?.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().ElementAt(sheetPosition - 1);
            if (sheet == null) throw new InvalidDataException("Invalid sheet");
            var workSheetPart = workbookPart.GetPartById(sheet.Id?.ToString() ?? string.Empty) as WorksheetPart;
            if (workSheetPart == null) throw new InvalidDataException("Invalid sheet");
            return workSheetPart;
        }

        WorksheetPart GetWorkSheetPartByName(string sheetName)
        {
            var workbookPart = _document.WorkbookPart;
            if (workbookPart == null) throw new InvalidDataException("Invalid sheet");
            var sheet = workbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>().FirstOrDefault(s => s.Name == sheetName);
            if (sheet == null) throw new InvalidDataException("Invalid sheet");
            var workSheetPart = workbookPart.GetPartById(sheet.Id?.ToString() ?? string.Empty) as WorksheetPart;
            if (workSheetPart == null) throw new InvalidDataException("Invalid sheet");
            return workSheetPart;
        }

        internal List<List<CellInfo>> GetCellInfo(int sheetPosition, double pdfWidthSrc, double pdfHeightSrc, out double scaling)
            => GetCellInfo(Workbook.Worksheet(sheetPosition), GetWorkSheetPartByPosition(sheetPosition), pdfWidthSrc, pdfHeightSrc, out scaling);

        internal List<List<CellInfo>> GetCellInfo(string sheetName, double pdfWidthSrc, double pdfHeightSrc, out double scaling)
            => GetCellInfo(Workbook.Worksheet(sheetName), GetWorkSheetPartByName(sheetName), pdfWidthSrc, pdfHeightSrc, out scaling);

        List<List<CellInfo>> GetCellInfo(IXLWorksheet ws, WorksheetPart worksheetPart, double pdfWidthSrc, double pdfHeightSrc, out double scaling)
            => GetCellInfo(ws.PageSetup, pdfWidthSrc, pdfHeightSrc,
                GetPageRanges(ws, worksheetPart), ws.MergedRanges.ToArray(), ws.Pictures.OfType<IXLPicture>().ToArray(), out scaling);

        internal XPen ConvertToXPen(XLBorderStyleValues borderStyle, XLColor? color, double scale)
        {
            var xcolor = ChangeColor(color) ?? XColor.FromArgb(255, 0, 0, 0);

            double lineWidth = 1.0;
            switch (borderStyle)
            {
                case XLBorderStyleValues.None:
                    lineWidth = 0;
                    break;
                case XLBorderStyleValues.Thin:
                    lineWidth = 0.5;
                    break;
                case XLBorderStyleValues.Medium:
                case XLBorderStyleValues.MediumDashDot:
                case XLBorderStyleValues.MediumDashDotDot:
                case XLBorderStyleValues.MediumDashed:
                    lineWidth = 1.5;
                    break;
                case XLBorderStyleValues.Thick:
                    lineWidth = 2.5;
                    break;
            }

            var pen = new XPen(xcolor, lineWidth * scale);
            switch (borderStyle)
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

        internal List<Color?> GetAccentColorsFromExcelTheme()
        {
            var workbookPart = _document.WorkbookPart;
            var themePart = workbookPart!.ThemePart;
            var theme = themePart!.Theme;
            var colorScheme = theme.ThemeElements!.ColorScheme;

            return new List<Color?>
            {
                ConvertToColor(colorScheme!.Accent1Color!),
                ConvertToColor(colorScheme.Accent2Color!),
                ConvertToColor(colorScheme.Accent3Color!),
                ConvertToColor(colorScheme.Accent4Color!),
                ConvertToColor(colorScheme.Accent5Color!),
                ConvertToColor(colorScheme.Accent6Color!)
            };
        }

        internal XColor? ChangeColor(XLColor? src)
        {
            if (src == null || !src.HasValue) return null;

            if (src.ColorType == XLColorType.Color)
            {
                if (src.Color.A == 0) return null;
                var colorValue = src.Color;
                return XColor.FromArgb(colorValue.A, colorValue.R, colorValue.G, colorValue.B);
            }
            else if (src.ColorType == XLColorType.Theme)
            {
                Color? netColor = null;
                if (XLThemeColor.Accent1 <= src.ThemeColor && src.ThemeColor <= XLThemeColor.Accent6)
                {
                    var list = GetAccentColorsFromExcelTheme();
                    netColor = list[src.ThemeColor - XLThemeColor.Accent1];

                }
                if (netColor == null)
                {
                    var resolvedColor1 = Workbook.Theme.ResolveThemeColor(src.ThemeColor);
                    var resolvedColor = resolvedColor1.Color;
                    netColor = Color.FromArgb(resolvedColor.A, resolvedColor.R, resolvedColor.G, resolvedColor.B);
                }
                var colorValue = ApplyTint(netColor.Value, src.ThemeTint);
                return XColor.FromArgb(colorValue.A, colorValue.R, colorValue.G, colorValue.B);

            }
            else if (src.ColorType == XLColorType.Indexed)
            {
                if (XLColor.IndexedColors.TryGetValue(src.Indexed, out var color) && color.ColorType == XLColorType.Color)
                {
                    var colorValue = color.Color;
                    if (colorValue.A == 0) return null;
                    return XColor.FromArgb(colorValue.A, colorValue.R, colorValue.G, colorValue.B);
                }
            }
            return null;
        }

        static Color? ConvertToColor(Color2Type dark1Color)
        {
            var rgbColorModelHex = dark1Color.RgbColorModelHex;
            if (rgbColorModelHex != null)
            {
                var rgb = rgbColorModelHex.Val!.Value;
                return Color.FromArgb(int.Parse(rgb!.Substring(0, 2), System.Globalization.NumberStyles.HexNumber),
                                      int.Parse(rgb.Substring(2, 2), System.Globalization.NumberStyles.HexNumber),
                                      int.Parse(rgb.Substring(4, 2), System.Globalization.NumberStyles.HexNumber));
            }

            var rgbColorModelPercentage = dark1Color.RgbColorModelPercentage;
            if (rgbColorModelPercentage != null)
            {
                byte r = ConvertPercentageToByteValue(rgbColorModelPercentage.RedPortion!);
                byte g = ConvertPercentageToByteValue(rgbColorModelPercentage.GreenPortion!);
                byte b = ConvertPercentageToByteValue(rgbColorModelPercentage.BluePortion!);
                return Color.FromArgb(r, g, b);
            }
            return null;
        }

        static byte ConvertPercentageToByteValue(Int32Value value)
            => (byte)(255 * value.Value / 100.0);

        static Color ApplyTint(Color originalColor, double tint)
        {
            if (tint == 0)
                return originalColor;

            double factor;

            if (tint < 0)
            {
                factor = (1.0 + tint) * 255;
                return Color.FromArgb(
                    originalColor.A,
                    (byte)(originalColor.R * factor / 255.0),
                    (byte)(originalColor.G * factor / 255.0),
                    (byte)(originalColor.B * factor / 255.0)
                );
            }
            else
            {
                factor = tint * (255.0 - originalColor.R) + originalColor.R;
                var r = (byte)factor;

                factor = tint * (255.0 - originalColor.G) + originalColor.G;
                var g = (byte)factor;

                factor = tint * (255.0 - originalColor.B) + originalColor.B;
                var b = (byte)factor;

                return Color.FromArgb(originalColor.A, r, g, b);
            }
        }

        class StartEnd
        {
            internal int Start { get; set; }
            internal int End { get; set; }
        }

        class Size
        {
            internal double Height { get; set; }
            internal double Width { get; set; }
        }

        internal void GetPageRanges(IXLWorksheet ws, int sheetPos, int? horizontalPageBreak = null, int? verticalPageBreak = null)
            =>GetPageRanges(ws, GetWorkSheetPartByPosition(sheetPos), horizontalPageBreak, verticalPageBreak);

        IXLRange[] GetPageRanges(IXLWorksheet ws, WorksheetPart worksheetPart, int? horizontalPageBreak = null, int? verticalPageBreak = null)
        {
            GetSheetMaxRowCol(worksheetPart, out var maxRow, out var maxColumn);
            if (maxRow == 0 || maxColumn == 0) return new IXLRange[0];

            var rowRanges = new List<StartEnd>();
            var rowIndex = 1;
            for (int i = 0; i < ws.PageSetup.RowBreaks.Count; i++)
            {
                var rowBreak = ws.PageSetup.RowBreaks[i];
                rowRanges.Add(new StartEnd { Start = rowIndex, End = rowBreak });
                rowIndex = rowBreak + 1;
            }
            if (rowIndex != maxRow)
            {
                if (rowIndex <= maxRow) rowRanges.Add(new StartEnd { Start = rowIndex, End = maxRow });
            }
            else if (!rowRanges.Any()) rowRanges.Add(new StartEnd { Start = rowIndex, End = maxRow });

            var colRanges = new List<StartEnd>();
            var colIndex = 1;
            for (int i = 0; i < ws.PageSetup.ColumnBreaks.Count; i++)
            {
                var colBreak = ws.PageSetup.ColumnBreaks[i];
                colRanges.Add(new StartEnd { Start = colIndex, End = colBreak });
                colIndex = colBreak + 1;
            }
            if (colIndex != maxColumn)
            {
                if (colIndex <= maxColumn) colRanges.Add(new StartEnd { Start = colIndex, End = maxColumn });
            }
            else if (!colRanges.Any()) colRanges.Add(new StartEnd { Start = 1, End = maxColumn });

            var list = new List<IXLRange>();
            foreach (var row in rowRanges)
            {
                foreach (var col in colRanges)
                {
                    list.Add(ws.Range(row.Start, col.Start, row.End, col.End));
                }
            }

            // Set the Break　page information (row and column)
            var rowBreaks = worksheetPart.Worksheet.Elements<RowBreaks>().FirstOrDefault();
            if(rowBreaks == null && horizontalPageBreak != null)
            {
                ws.PageSetup.AddHorizontalPageBreak(horizontalPageBreak ?? 0);
            }

            var colBreaks = worksheetPart.Worksheet.Elements<ColumnBreaks>().FirstOrDefault();
            if (rowBreaks == null && verticalPageBreak != null)
            {
                ws.PageSetup.AddVerticalPageBreak(verticalPageBreak ?? 0);
            }

            return list.ToArray();
        }

        internal void GetSheetMaxRowCol(int sheetPos, out int maxRow, out int maxColumn)
            => GetSheetMaxRowCol(GetWorkSheetPartByPosition(sheetPos), out maxRow, out maxColumn);

        void GetSheetMaxRowCol(WorksheetPart worksheetPart, out int maxRow, out int maxColumn)
        {
            // 1. Enumerate all rows and cells
            var rows = worksheetPart.Worksheet.Descendants<Row>();

            if (!rows.Any())
            {
                maxRow = 0;
                maxColumn = 0;
                return;
            }

            // 2. Calculate the minimum and maximum range
            uint uintMaxRow = rows.Max(r => r.RowIndex) ?? 0;
            maxRow = (int)uintMaxRow;

            maxColumn = int.MinValue;
            foreach (var row in rows)
            {
                var cells = row.Elements<Cell>();
                foreach (var cell in cells)
                {
                    var cellReference = cell.CellReference;
                    if (!string.IsNullOrEmpty(cell.CellReference))
                    {
                        int columnIndex = GetColumnIndex(cell.CellReference);
                        if (columnIndex > maxColumn) maxColumn = columnIndex;
                    }
                }
            }

            // 3. If the column does not exist
            if (maxColumn == int.MinValue)
            {
                maxColumn = 1; // Column A
            }
        }

        // Get column number from cell reference
        static int GetColumnIndex(string? cellReference)
        {
            if (string.IsNullOrEmpty(cellReference))
            {
                return 0;
            }

            var colPart = new string(cellReference.Where(char.IsLetter).ToArray());
            int colIndex = 0;
            foreach (char c in colPart)
            {
                colIndex = (colIndex * 26) + (c - 'A' + 1);
            }
            return colIndex;
        }

        // Get column name from column number
        static string GetColumnName(int columnIndex)
        {
            var columnName = string.Empty;
            while (columnIndex > 0)
            {
                var remainder = (columnIndex - 1) % 26;
                columnName = (char)(remainder + 'A') + columnName;
                columnIndex = (columnIndex - remainder - 1) / 26;
            }
            return columnName;
        }

        internal static double PixelToPoint(double src)
            => src * (72.0 / 96.0);

        static double InchToPoint(double inch)
            => inch * 72;

        static double ColumnWidthToPoint(double columnWidth)
            => PixelToPoint(ColumnWithToPixel(columnWidth));

        static double ColumnWithToPixel(double columnWidth)
            => columnWidth * 8.0 + 5.0;

        static List<List<CellInfo>> GetCellInfo(
            IXLPageSetup pageSetup, double pdfWidthSrc, double pdfHeightSrc,
            IXLRange[] ranges, IXLRange[] mergedRanges, IXLPicture[] pictures, out double scaling)
        {
            var indexAndPictures = pictures.Select((Picture, Index) => new { Picture, Index }).ToList();

            scaling = ((double)pageSetup.Scale) / 100;

            var allCells = new List<List<CellInfo>>();
            foreach (var range in ranges)
            {
                var (marginX, marginY) = GetMargin(pageSetup, pdfWidthSrc, pdfHeightSrc, range);

                double yOffset = 0;
                var cells = new List<CellInfo>();

                foreach (var row in range.Rows())
                {
                    double xOffset = 0;
                    double scaledHeight = 0;
                    foreach (var cell in row.Cells())
                    {
                        var cellWidth = ColumnWidthToPoint(cell.WorksheetColumn().Width);
                        var cellHeight = cell.WorksheetRow().Height;

                        // Calculate scaling
                        var scaledWidth = cellWidth * scaling;
                        scaledHeight = cellHeight * scaling;

                        // CellInfo
                        var info = new CellInfo
                        {
                            X = xOffset + marginX,
                            Y = yOffset + marginY,
                            Width = scaledWidth,
                            Height = scaledHeight,
                            Cell = cell
                        };

                        //Add Picture to Cell
                        foreach (var e in indexAndPictures.Where(e => e.Picture.TopLeftCell.Address.UniqueId == cell.Address.UniqueId))
                        {
                            e.Picture.ImageStream.Position = 0;
                            info.Pictures.Add(new()
                            {
                                Picture = e.Picture.ImageStream,
                                Index = e.Index,
                                X = PixelToPoint(e.Picture.Left) * scaling,
                                Y = PixelToPoint(e.Picture.Top) * scaling,
                                Width = PixelToPoint(e.Picture.Width) * scaling,
                                Height = PixelToPoint(e.Picture.Height) * scaling
                            });
                        }

                        cells.Add(info);

                        // Add xOffset
                        xOffset += scaledWidth;
                    }

                    // Add yOffset
                    yOffset += scaledHeight;
                }

                allCells.Add(cells);
            }

            // Add margin info
            var infoList = allCells.SelectMany(e => e);
            foreach (var range in mergedRanges)
            {
                var firstCellId = range.FirstCell().Address.UniqueId;
                var lastCellId = range.LastCell().Address.UniqueId;
                var firstInfo = infoList.FirstOrDefault(e => e.Cell?.Address.UniqueId == firstCellId);
                var lastInfo = infoList.FirstOrDefault(e => e.Cell?.Address.UniqueId == lastCellId);
                if (firstInfo == null) continue;
                double w = 0, h = 0;
                bool getW = true;
                foreach (var row in range.Rows())
                {
                    if (getW)
                    {
                        foreach (var cell in row.Cells())
                        {
                            w += cell.WorksheetColumn().Width;
                        }
                    }
                    getW = false;
                    h += row.WorksheetRow().Height;
                }
                firstInfo.MergedWidth = ColumnWidthToPoint(w) * scaling;
                firstInfo.MergedHeight = h * scaling;

                foreach (var row in range.Rows())
                {
                    foreach (var cell in row.Cells())
                    {
                        var merged = infoList.FirstOrDefault(e => e.Cell?.Address.UniqueId == cell.Address.UniqueId);
                        if (merged == null) continue;
                        merged.MergedFirstCell = firstInfo;
                        merged.MergedLastCell = lastInfo;
                    }
                }
            }
            return allCells;
        }

        static (double marginX, double marginY) GetMargin(IXLPageSetup pageSetup, double pdfWidthSrc, double pdfHeightSrc, IXLRange range)
        {
            var marginLeft = InchToPoint(pageSetup.Margins.Left);
            var marginTop = InchToPoint(pageSetup.Margins.Top + pageSetup.Margins.Header);

            if (!pageSetup.CenterHorizontally && !pageSetup.CenterVertically) return (marginLeft, marginTop);

            var marginRight = InchToPoint(pageSetup.Margins.Right);
            var marginBottom = InchToPoint(pageSetup.Margins.Bottom + pageSetup.Margins.Footer);

            double totalWidth = 0;
            double totalHeight = 0;
            foreach (var row in range.Rows())
            {
                totalHeight += row.WorksheetRow().Height;
                if (totalWidth == 0)
                {
                    foreach (var cell in row.Cells())
                    {
                        totalWidth += ColumnWidthToPoint(cell.WorksheetColumn().Width);
                    }
                }
            }

            var marginX = marginLeft;
            var marginY = marginTop;
            var pdfWidth = pdfWidthSrc - marginX - marginRight;
            var pdfHeight = pdfHeightSrc - marginY - marginBottom;

            if (pageSetup.CenterHorizontally)
            {
                if (totalWidth < pdfWidth)
                {
                    marginX += ((pdfWidth - totalWidth) / 2);
                }
            }
            if (pageSetup.CenterVertically)
            {
                if (totalHeight < pdfHeight)
                {
                    marginY += ((pdfHeight - totalHeight) / 2);
                }
            }

            return (marginX, marginY);
        }
    }
}
