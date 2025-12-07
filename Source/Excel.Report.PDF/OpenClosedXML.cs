using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Color = System.Drawing.Color;

namespace Excel.Report.PDF
{
    class Margins
    {
        internal double Left { get; set; }
        internal double Right { get; set; }
        internal double Top { get; set; }
        internal double Bottom { get; set; }
        internal double Header { get; set; }
        internal double Footer { get; set; }
    }

    class PageSetup
    {
        internal Margins Margins { get; set; } = new Margins();
        internal bool CenterHorizontally { get; set; }
        internal bool CenterVertically { get; set; }
        internal int Scale { get; set; } = 100;
        internal double Width { get; set; }
        internal double Height { get; set; }
        internal bool IsFitColumn { get; set; }

        internal static PageSetup FromIXLPageSetup(IXLPageSetup pageSetup)
        {
            (var w, var h) = PaperSizeMap.GetPaperSize(pageSetup.PaperSize);
            return new PageSetup
            {
                Margins = new Margins
                {
                    Left = pageSetup.Margins.Left,
                    Right = pageSetup.Margins.Right,
                    Top = pageSetup.Margins.Top,
                    Bottom = pageSetup.Margins.Bottom,
                    Header = pageSetup.Margins.Header,
                    Footer = pageSetup.Margins.Footer
                },
                CenterHorizontally = pageSetup.CenterHorizontally,
                CenterVertically = pageSetup.CenterVertically,
                Scale = pageSetup.Scale,
                Width = w.Point,
                Height = h.Point,
                IsFitColumn = 0 < pageSetup.PagesWide
            };
        }
    }

    class RenderInfo
    {
        internal List<CellInfo> Cells { get; set; } = new List<CellInfo>();
        internal double Scaling { get; set; }
    }

    class OpenClosedXML : IDisposable
    {
        readonly SpreadsheetDocument _document; 

        internal XLWorkbook Workbook { get; }

        internal IXLPageSetup GetPageSetup(int sheetPosition)
            => Workbook.Worksheet(sheetPosition).PageSetup;

        internal int GetSheetPosition(string sheetName)
            => Workbook.Worksheet(sheetName)?.Position ?? -1;

        internal int SheetCount
            => Workbook.Worksheets.Count;

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
            var sheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().ElementAt(sheetPosition - 1);
            if (sheet == null) throw new InvalidDataException("Invalid sheet");
            var workSheetPart = workbookPart.GetPartById(sheet.Id?.ToString() ?? string.Empty) as WorksheetPart;
            if (workSheetPart == null) throw new InvalidDataException("Invalid sheet");
            return workSheetPart;
        }

        internal List<RenderInfo> GetCellInfo(PageSetup pageSetup, int sheetPosition)
        {
            var ws = Workbook.Worksheet(sheetPosition);
            var worksheetPart = GetWorkSheetPartByPosition(sheetPosition);
            var ranges = GetPageRanges(ws, worksheetPart);

            var text = ws.GetText(1, 1);
            var specialKeys = text.Split('|').Select(e => e.Trim()).ToList();
            var isFitColumn = specialKeys.Contains("#FitColumn");

            return GetCellInfo(pageSetup, ranges, 
                ws.MergedRanges.ToArray(), ws.Pictures.OfType<IXLPicture>().ToArray(), isFitColumn);
        }

        internal VirtualPen ConvertToPen(XLBorderStyleValues borderStyle, XLColor? color, double scale)
        {
            var xcolor = ChangeColor(color) ?? new VirtualColor(255, 0, 0, 0);

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

            var pen = new VirtualPen(xcolor, lineWidth * scale, borderStyle);
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

        internal VirtualColor? ChangeColor(XLColor? src)
        {
            if (src == null || !src.HasValue) return null;

            if (src.ColorType == XLColorType.Color)
            {
                if (src.Color.A == 0) return null;
                var colorValue = src.Color;
                return new VirtualColor(colorValue.A, colorValue.R, colorValue.G, colorValue.B);
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
                return new VirtualColor(colorValue.A, colorValue.R, colorValue.G, colorValue.B);

            }
            else if (src.ColorType == XLColorType.Indexed)
            {
                if (XLColor.IndexedColors.TryGetValue(src.Indexed, out var color) && color.ColorType == XLColorType.Color)
                {
                    var colorValue = color.Color;
                    if (colorValue.A == 0) return null;
                    return new VirtualColor(colorValue.A, colorValue.R, colorValue.G, colorValue.B);
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

        internal void GetPageRanges(IXLWorksheet ws, int sheetPos)
            =>GetPageRanges(ws, GetWorkSheetPartByPosition(sheetPos));

        IXLRange[] GetPageRanges(IXLWorksheet ws, WorksheetPart worksheetPart)
        {
            GetSheetMaxRowCol(ws, worksheetPart, out var maxRow, out var maxColumn);
            if (maxRow == 0 || maxColumn == 0) return new IXLRange[0];

            return GetPageRangesByExcelOrder(ws, maxRow, maxColumn);
        }
        
        static IXLRange[] GetPageRangesByExcelOrder(IXLWorksheet ws, int maxRow, int maxColumn)
        {
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
            return list.ToArray();
        }

        internal void GetSheetMaxRowCol(int sheetPos, out int maxRow, out int maxColumn)
            => GetSheetMaxRowCol(Workbook.Worksheet(sheetPos), GetWorkSheetPartByPosition(sheetPos), out maxRow, out maxColumn);

        void GetSheetMaxRowCol(IXLWorksheet ws, WorksheetPart worksheetPart, out int maxRow, out int maxColumn)
        {
            var picPositions = ws.Pictures
                .OfType<IXLPicture>()
                .Select(p => new
                {
                    Row = p.TopLeftCell.Address.RowNumber,
                    Column = p.TopLeftCell.Address.ColumnNumber
                })
                .ToList();
            maxRow = picPositions.Any() ? picPositions.Max(e=>e.Row) : 1;
            maxColumn = picPositions.Any() ? picPositions.Max(e => e.Column) : 1;

            // 1. Enumerate all rows and cells
            var rows = worksheetPart.Worksheet.Descendants<Row>();

            if (!rows.Any())
            {
                return;
            }

            // 2. Calculate the minimum and maximum range
            uint uintMaxRow = rows.Max(r => r.RowIndex) ?? 0;
            maxRow = Math.Max(maxRow, (int)uintMaxRow);

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

        internal static double PixelToPoint(double src)
            => src * (72.0 / 96.0);

        static double InchToPoint(double inch)
            => inch * 72;

        static double ColumnWidthToPoint(double columnWidth)
            => PixelToPoint(ColumnWidthToPixel(columnWidth));

        static double ColumnWidthToPixel(double columnWidth)
        {
            // Excel Maximum Digit Width (MDW). For your case (aiming for 97 px), MDW = 8 fits.
            // If it comes out 1–2 px smaller with the default Calibri 11, change it to 7 and verify.
            const double mdw = 8;

            // Left/right cell padding: 2*CEILING(MDW/4)+1
            double pp = 2 * (double)Math.Ceiling(mdw / 4.0) + 1;

            if (columnWidth < 1.0)
            {
                // Nonlinear region for NoC < 1
                return columnWidth * (mdw + pp);
            }

            // Excel-compliant: round to 1/256 of a character, multiply by MDW,
            // then add 0.5 and round
            double noc256 = (256.0 * columnWidth + Math.Round(128.0 / mdw)) / 256.0;
            return noc256 * mdw + pp;
        }

        static List<RenderInfo> GetCellInfo(
            PageSetup pageSetup, IXLRange[] ranges, IXLRange[] mergedRanges, IXLPicture[] pictures, bool isFitColumn)
        {
            var indexAndPictures = pictures.Select((Picture, Index) => new { Picture, Index }).ToList();

            var renderInfoList = new List<RenderInfo>();
            foreach (var range in ranges)
            {
                var (marginX, marginY, scaling) = GetMarginAndScaling(pageSetup, range, isFitColumn);

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

                renderInfoList.Add(new RenderInfo { Cells = cells, Scaling = scaling });
            }

            foreach(var renderInfo in renderInfoList)
            {            
                // Add margin info
                var infoList = renderInfo.Cells.ToList();
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
                                w += ColumnWidthToPoint(cell.WorksheetColumn().Width);
                            }
                        }
                        getW = false;
                        h += row.WorksheetRow().Height;
                    }
                    firstInfo.MergedWidth = w * renderInfo.Scaling;
                    firstInfo.MergedHeight = h * renderInfo.Scaling;

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
            }

            return renderInfoList;
        }

        static (double marginX, double marginY, double scaling) GetMarginAndScaling(PageSetup pageSetup, IXLRange range, bool isFitColumn)
        {
            var marginLeft = InchToPoint(pageSetup.Margins.Left);
            var marginTop = InchToPoint(pageSetup.Margins.Top + pageSetup.Margins.Header);


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

            var scaling = ((double)pageSetup.Scale) / 100;
            if (scaling == 0) scaling = 1.0;

            //FitColumn
            if (isFitColumn || pageSetup.IsFitColumn)
            {
                var pdfWidth = pageSetup.Width - marginLeft - marginRight;
                scaling = pdfWidth / totalWidth;
                totalWidth = totalWidth * scaling;
                totalHeight = totalHeight * scaling;
            }

            var marginX = marginLeft;
            var marginY = marginTop;

            if (pageSetup.CenterHorizontally)
            {
                var pdfWidth = pageSetup.Width - marginX - marginRight;
                if (totalWidth < pdfWidth)
                {
                    marginX += ((pdfWidth - totalWidth) / 2);
                }
            }
            if (pageSetup.CenterVertically)
            {
                var pdfHeight = pageSetup.Height - marginY - marginBottom;
                if (totalHeight < pdfHeight)
                {
                    marginY += ((pdfHeight - totalHeight) / 2);
                }
            }

            return (marginX, marginY, scaling);
        }

        internal List<string> GetSheetNames()
        {
            var sheetNames = Workbook.Worksheets.OrderBy(e => e.Position).Select(e => e.Name).ToList();
            return sheetNames;
        }
    }
}
