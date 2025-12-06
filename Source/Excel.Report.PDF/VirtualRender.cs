using ClosedXML.Excel;

namespace Excel.Report.PDF
{
    class VirtualRender
    {
        class PageProcessCommand : IPostProcessCommand
        {
            Action Action { get; }
            public PageProcessCommand(Action action) => Action = action;
            public void Execute() => Action();
        }

        readonly OpenClosedXML _openClosedXML;
        readonly List<IPostProcessCommand> _postProcessCommands = new();

        internal VirtualRender(OpenClosedXML openClosedXML)
            => _openClosedXML = openClosedXML;

        internal void RenderTo(IVirtualDocument document, PageSetup? pageSetup)
        {
            for (int i = 1; i <= _openClosedXML.SheetCount; i++)
            {
                RenderCore(document, pageSetup, i);
            }
            _postProcessCommands.ExecuteAll();
        }

        internal void RenderTo(IVirtualDocument document, PageSetup? pageSetup, int sheetPosition)
        {
            RenderCore(document, pageSetup, sheetPosition);
            _postProcessCommands.ExecuteAll();
        }

        void RenderCore(IVirtualDocument document, PageSetup? pageSetup, int sheetPosition)
        {
            var ps = _openClosedXML.GetPageSetup(sheetPosition);
            var ws = _openClosedXML.Workbook.Worksheet(sheetPosition);
            if (pageSetup == null) pageSetup = PageSetup.FromIXLPageSetup(ws.PageSetup);
            var allCells = _openClosedXML.GetCellInfo(pageSetup, sheetPosition);
            RenderTo(document, ps, allCells);
        }

        void RenderTo(IVirtualDocument document, IXLPageSetup ps, List<RenderInfo> boolRenderInfo)
        {
            foreach(var sheetRenderInfo in boolRenderInfo)
            {
                var page = document.AddPage(ps);
                var gfx = page.CreateGraphics();
                var drawLineCache = new DrawLineCache(gfx);

                // Since there are duplicate parts, the loops are separated to prevent overwriting.
                foreach (var cellInfo in sheetRenderInfo.Cells)
                {
                    FillCellBackColor(gfx, cellInfo);
                }
                foreach (var cellInfo in sheetRenderInfo.Cells)
                {
                    DrawRuledLine(drawLineCache, sheetRenderInfo.Scaling, cellInfo);
                }
                foreach (var cellInfo in sheetRenderInfo.Cells)
                {
                    DrawCellText(document, gfx, sheetRenderInfo.Scaling, cellInfo);
                }

                var pictureInfoAndCellInfo = new List<PictureInfoAndCellInfo>();
                foreach (var cellInfo in sheetRenderInfo.Cells)
                {
                    foreach (var e in cellInfo.Pictures)
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

        void FillCellBackColor(IVirtualGraphics gfx, CellInfo cellInfo)
        {
            var cell = cellInfo.Cell!;
            if (cellInfo.MergedFirstCell != null) cell = cellInfo.MergedFirstCell.Cell!;

            var xBackColor = _openClosedXML.ChangeColor(cell.Style.Fill.BackgroundColor);
            if (xBackColor != null)
            {
                gfx.DrawRectangle(xBackColor, cellInfo.X, cellInfo.Y, cellInfo.Width, cellInfo.Height);
            }
        }

        // If you draw two lines in the same place, it will be darker, so skip the second one.
        class DrawLineCache
        {
            IVirtualGraphics _gfx;
            Dictionary<string, bool> _cache = new Dictionary<string, bool>();
            public DrawLineCache(IVirtualGraphics gfx) => _gfx = gfx;

            public void DrawLine(VirtualPen xPen, double x1, double y1, double x2, double y2)
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
                    var thin = _openClosedXML.ConvertToPen(XLBorderStyleValues.Thin, color, scaling);

                    // Ensure a visible gap on rasterizers
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
                var pen = _openClosedXML.ConvertToPen(style, color, scaling);
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

        void DrawCellText(IVirtualDocument document, IVirtualGraphics currentXG, double scaling, CellInfo cellInfo)
        {
            var cell = cellInfo.Cell!;
            var text = cell.GetFormattedString();
            var specialKeys = text.Split('|').Select(e => e.Trim()).ToList();
            if (specialKeys.Contains("#Empty")) return;
            if (specialKeys.Contains("#FitColumn")) return;

            //special formating: do not draw
            if (cell.Style.NumberFormat.Format == ";;;") return;

            // Alignment
            var format = new VirtualStringFormat();
            switch (cell.Style.Alignment.Horizontal)
            {
                case XLAlignmentHorizontalValues.Center:
                    format.Alignment = VirtualAlignment.Center;
                    break;
                case XLAlignmentHorizontalValues.Right:
                    format.Alignment = VirtualAlignment.Far;
                    break;
                default:
                    switch (cell.DataType)
                    {
                        case XLDataType.Number:
                        case XLDataType.DateTime:
                            format.Alignment = VirtualAlignment.Far;
                            break;
                        case XLDataType.Boolean:
                            format.Alignment = VirtualAlignment.Center;
                            break;
                        default:
                            format.Alignment = VirtualAlignment.Near;
                            break;
                    }
                    break;
            }
            switch (cell.Style.Alignment.Vertical)
            {
                case XLAlignmentVerticalValues.Center:
                    format.LineAlignment = VirtualAlignment.Center;
                    break;
                case XLAlignmentVerticalValues.Bottom:
                    format.LineAlignment = VirtualAlignment.Far;
                    break;
                default:
                    format.LineAlignment = VirtualAlignment.Near;
                    break;
            }

            // Font
            double fontSize = cell.Style.Font.FontSize;
            var font = new VirtualFont(cell.Style.Font, scaling);
            var xFontColor = _openClosedXML.ChangeColor(cell.Style.Font.FontColor) ?? new VirtualColor(255, 0, 0, 0);

            double w = cellInfo.MergedWidth != 0 ? cellInfo.MergedWidth : cellInfo.Width;
            double h = cellInfo.MergedHeight != 0 ? cellInfo.MergedHeight : cellInfo.Height;

            // Excel-like padding
            var cellPaddingPt = OpenClosedXML.PixelToPoint(fontSize * (1.0 / 4.0));
            var offset = cellPaddingPt * scaling;
            if (offset * 2 < w) w -= offset * 2;
            if (offset * 2 < h) h -= offset * 2;

            var rect = new VirtualRect(cellInfo.X + offset, cellInfo.Y + offset, w, h);

            var lines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

            // ===== Rotation & vertical text =====
            int raw = cell.Style.Alignment.TextRotation;

            if (raw == 255)
            {
                // Excel's "Vertical Text" (stack)
                if (TryResolvePageVariable(lines, l => DrawVerticalStack(currentXG, font, xFontColor, rect, format, new[] { l }))) return;
                DrawVerticalStack(currentXG, font, xFontColor, rect, format, lines);
                return;
            }

            // Excel (0..90 = counterclockwise / 91..180 = clockwise (= negative angle))
            int angle = 0;
            if (raw <= 90) angle = -raw;        // Up-left slant (Excel +) → negative angle in target
            else angle = 180 - raw;    // Up-right slant (Excel -) → positive angle in target

            if (angle != 0)
            {
                if (TryResolvePageVariable(lines, l => DrawRotated(currentXG, font, xFontColor, rect, format, new[] { l }, angle))) return;
                DrawRotated(currentXG, font, xFontColor, rect, format, lines, angle);
                return;
            }

            var fontHeight = currentXG.GetFontHeight(font);

            // ===== Horizontal text (no rotation) =====
            double startY = rect.Y;
            if (format.LineAlignment == VirtualAlignment.Center)
                startY += (rect.Height - lines.Length * fontHeight) / 2.0;
            else if (format.LineAlignment == VirtualAlignment.Far)
                startY += (rect.Height - lines.Length * fontHeight);

            if (TryResolvePageVariable(lines, l => currentXG.DrawString(l, font, xFontColor, new VirtualRect(rect.X, startY, rect.Width, fontHeight), format))) return;
            foreach (var line in lines)
            {
                currentXG.DrawString(line, font, xFontColor, new VirtualRect(rect.X, startY, rect.Width, fontHeight), format);
                startY += fontHeight;
            }

            // ======== Local functions ========
            bool TryResolvePageVariable(string[] lines, Action<string> draw)
            {
                if (lines.Length != 1) return false;
                var line = lines[0];

                Action action = () => { };
                if (line == "#Page")
                {
                    line = document.PageCount.ToString();
                    draw(line);
                    return true;
                }
                else if (line == "#PageCount")
                {
                    _postProcessCommands.Add(new PageProcessCommand(() =>
                    {
                        var pageCount = document.PageCount.ToString();
                        draw(pageCount);
                    }));
                    return true;
                }
                else if (line.StartsWith("#PageOf"))
                {
                    var args = line.Replace("#PageOf", "").Replace("(", "").Replace(")", "").Split(',').Select(e => e.Trim()).ToArray();
                    var sp = args.FirstOrDefault()?.Replace("\"", "") ?? "/";
                    var currentPage = document.PageCount.ToString();
                    _postProcessCommands.Add(new PageProcessCommand(() =>
                    {
                        var pageCount = document.PageCount.ToString();
                        draw(currentPage + sp + pageCount);
                    }));
                    return true;
                }
                return false;
            }

            // Vertical writing (Excel stack): place characters top→bottom, advance columns left→right
            static void DrawVerticalStack(IVirtualGraphics g, VirtualFont f, VirtualColor b, VirtualRect r, VirtualStringFormat fmt, string[] cols)
            {
                double step = g.GetFontHeight(f);                 // one cell
                double totalW = cols.Length * step;

                double startX = r.X;
                if (fmt.Alignment == VirtualAlignment.Center)
                    startX += Math.Max(0, (r.Width - totalW) / 2.0);
                else if (fmt.Alignment == VirtualAlignment.Far)
                    startX += Math.Max(0, r.Width - totalW);

                var charFmt = new VirtualStringFormat { Alignment = VirtualAlignment.Center, LineAlignment = VirtualAlignment.Near };

                for (int c = 0; c < cols.Length; c++)
                {
                    string col = cols[c] ?? string.Empty;
                    double colH = col.Length * step;

                    double y = r.Y;
                    if (fmt.LineAlignment == VirtualAlignment.Center)
                        y += Math.Max(0, (r.Height - colH) / 2.0);
                    else if (fmt.LineAlignment == VirtualAlignment.Far)
                        y += Math.Max(0, r.Height - colH);

                    double x = startX + c * step;

                    for (int i = 0; i < col.Length; i++)
                    {
                        string ch = col.Substring(i, 1);
                        g.DrawString(ch, f, b, new VirtualRect(x, y + i * step, step, step), charFmt);
                    }
                }
            }

            // Arbitrary-angle drawing: rotate the coordinate system around the rectangle center (do not swap width/height)
            static void DrawRotated(IVirtualGraphics g, VirtualFont f, VirtualColor b, VirtualRect r, VirtualStringFormat fmt, string[] content, int angle)
            {
                var fontHeight = g.GetFontHeight(f);

                g.Save();

                // Rotate about the center
                g.TranslateTransform(r.X + r.Width / 2.0, r.Y + r.Height / 2.0);
                g.RotateTransform(angle);

                var rr = new VirtualRect(-r.Width / 2.0, -r.Height / 2.0, r.Width, r.Height);

                double y = rr.Y;
                if (fmt.LineAlignment == VirtualAlignment.Center)
                    y += (rr.Height - content.Length * fontHeight) / 2.0;
                else if (fmt.LineAlignment == VirtualAlignment.Far)
                    y += rr.Height - content.Length * fontHeight;

                foreach (var line in content)
                {
                    g.DrawString(line, f, b, new VirtualRect(rr.X, y, rr.Width, fontHeight), fmt);
                    y += fontHeight;
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

        static void DrawPictures(IVirtualGraphics gfx, PictureInfoAndCellInfo item)
        {
            item.PictureInfo.Picture!.Position = 0;
            gfx.DrawImage(item.PictureInfo.Picture!,
            item.CellInfo.X + item.PictureInfo.X,
                item.CellInfo.Y + item.PictureInfo.Y,
                item.PictureInfo.Width,
                item.PictureInfo.Height);
        }
    }
}
