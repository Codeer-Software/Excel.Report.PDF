using ClosedXML.Excel;
using System.Collections;

namespace Excel.Report.PDF
{
    public static class ExcelOverWriter
    {
        class PageLoopRowsInfo
        { 
            public List<object?> List { get; set; } = new();

            public string FirstPageSheetName { get; set; } = string.Empty;
            public int FirstPageBlockCount { get; set; }
            public string SourceBodyPageSheetName { get; set; } = string.Empty;
            public List<string> BodyPageSheetNames { get; set; } = new();
            public int BodyPageBlockCount { get; set; }
            public string LastPageSheetName { get; set; } = string.Empty;
            public int LastPageBlockCount { get; set; }

            public List<object?> FirstPageList { get; set; } = new();
            public List<List<object?>> BodyPageLists { get; set; } = new();
            public List<object?> LastPageList { get; set; } = new();
        }

        enum PageType
        {
            First,
            Body,
            Last
        }

        public static async Task OverWrite(this XLWorkbook book, IExcelSymbolConverter converter)
        {
            var pagedLoopRowsInfos = await GetPagedLoopRowInfos(book, converter);

            AdjustBodySheets(book, pagedLoopRowsInfos);

            foreach (var sheet in book.Worksheets)
            {
                await OverWrite(sheet, converter, pagedLoopRowsInfos.Values.ToList());
            }
        }

        public static async Task OverWrite(this IXLWorksheet sheet, IExcelSymbolConverter converter)
            => await OverWrite(sheet, converter, new());

        static void AdjustBodySheets(XLWorkbook book, Dictionary<string, PageLoopRowsInfo> pagedLoopRowsInfos)
        {
            foreach (var e in pagedLoopRowsInfos)
            {
                if (string.IsNullOrEmpty(e.Value.SourceBodyPageSheetName)) continue;
                var bodySheet = book.Worksheet(e.Value.SourceBodyPageSheetName);
                for (int i = 0; i < e.Value.BodyPageLists.Count; i++)
                {
                    bodySheet.CopyTo($"{e.Value.SourceBodyPageSheetName}_{i}", bodySheet.Position + i);
                }
                book.Worksheets.Delete(e.Value.SourceBodyPageSheetName);
            }
        }

        static async Task<Dictionary<string, PageLoopRowsInfo>> GetPagedLoopRowInfos(XLWorkbook book, IExcelSymbolConverter converter)
        {
            Dictionary<string, PageLoopRowsInfo> pagedLoopRowsInfos = new();

            foreach (var sheet in book.Worksheets)
            {
                ExcelUtils.GetRowColCount(sheet, out var rowCount, out var colCount);

                // get left cells and check #PagedLoopRows
                List<string> leftCells = new();
                int pagedLoopCount = 0;
                for (int i = 0; i <= rowCount; i++)
                {
                    var rowIndex = i + 1;
                    var text = sheet.GetText(rowIndex, 1).Trim();
                    if (text.StartsWith("#PagedLoopRows")) pagedLoopCount++;
                    if (1 < pagedLoopCount) throw new Exception($"One sheet can have only one #PagedLoopRows. SheetName:{sheet.Name}");
                    leftCells.Add(text);
                }
                
                foreach(var leftCell in leftCells)
                {
                    //#PagedLoopRows(pageType, rowsPerPage, $items, items, blockRowCount)
                    if (leftCell.StartsWith("#PagedLoopRows"))
                    {
                        var args = leftCell.Replace("#PagedLoopRows", "").Replace("(", "").Replace(")", "").Split(',').Select(e => e.Trim()).ToArray();
                        if (args.Length != 5) break;
                        var items = args[2];
                        if (!items.StartsWith("$")) break;
                        items = items.Substring(1);
                        if (!pagedLoopRowsInfos.TryGetValue(items, out var info))
                        {
                            info = new PageLoopRowsInfo();
                            var enumerable = (await converter.GetData(items))?.Value as IEnumerable;
                            if (enumerable == null) break;
                            foreach (var e in enumerable)
                            {
                                info.List.Add(e);
                            }
                            pagedLoopRowsInfos[items] = info;
                        }
                        if (!Enum.TryParse<PageType>(args[0], out var pageType)) break;
                        if (!int.TryParse(args[1], out var rowsPerPage)) break;
                        if (!int.TryParse(args[4], out var blockRowCount)) break;
                        switch (pageType)
                        {
                            case PageType.First:
                                info.FirstPageSheetName = sheet.Name;
                                info.FirstPageBlockCount = rowsPerPage;
                                break;
                            case PageType.Body:
                                info.SourceBodyPageSheetName = sheet.Name;
                                info.BodyPageBlockCount = rowsPerPage;
                                break;
                            case PageType.Last:
                                info.LastPageSheetName = sheet.Name;
                                info.LastPageBlockCount = rowsPerPage;
                                break;
                        }
                        break;
                    }
                }
            }

            //Distributing Lists per page
            foreach (var e in pagedLoopRowsInfos)
            {
                if (!e.Value.List.Any())
                {
                    continue;
                }
                int bodyCount = e.Value.List.Count - e.Value.FirstPageBlockCount - e.Value.LastPageBlockCount;
                if (bodyCount == 0)
                {
                    var firstPageCount = e.Value.FirstPageBlockCount;
                    var lastPageCount = e.Value.List.Count - e.Value.FirstPageBlockCount;
                    if (lastPageCount <= 0)
                    {
                        lastPageCount = 1;
                        firstPageCount = e.Value.List.Count - 1;
                    }
                    e.Value.FirstPageList = e.Value.List.Take(firstPageCount).ToList();
                    e.Value.LastPageList = e.Value.List.Skip(firstPageCount).Take(lastPageCount).ToList();
                }
                else
                {
                    var first = e.Value.List.Take(e.Value.FirstPageBlockCount).ToList();
                    var body = new List<List<object?>>();
                    var rest = e.Value.List.Skip(e.Value.FirstPageBlockCount).ToList();
                    while (e.Value.LastPageBlockCount < rest.Count)
                    {
                        body.Add(rest.Take(e.Value.BodyPageBlockCount).ToList());
                        rest = rest.Skip(e.Value.BodyPageBlockCount).ToList();
                    }

                    e.Value.FirstPageList = first;
                    e.Value.BodyPageLists = body;
                    e.Value.LastPageList = rest;

                    for(int i = 0; i < body.Count; i++)
                    {
                        e.Value.BodyPageSheetNames.Add($"{e.Value.SourceBodyPageSheetName}_{i}");
                    }
                }
            }
            return pagedLoopRowsInfos;
        }

        static async Task OverWrite(this IXLWorksheet sheet, IExcelSymbolConverter converter, List<PageLoopRowsInfo> pageLoopRowsInfoList)
        {
            // Get all rows and columns of the sheet
            ExcelUtils.GetRowColCount(sheet, out var rowCount, out var colCount);
            await OverWrite(sheet, 1, rowCount, colCount, converter, pageLoopRowsInfoList);
        }

        static async Task<int> OverWrite(IXLWorksheet sheet, int startRow, int endRow, int colCount, IExcelSymbolConverter converter, List<PageLoopRowsInfo> pageLoopRowsInfoList)
        {
            for (int i = startRow; i <= endRow;)
            {
                await OverWriteCell(sheet, i, colCount, async t => await converter.GetData(t));

                LoopInfo loopInfo = new();
                if (!await TryParseLoop(sheet.GetText(i, 1).Trim(), converter, loopInfo, sheet.Name, pageLoopRowsInfoList))
                {
                    i++;
                    continue;
                }

                // delete #LoopRow
                var cell = sheet.Cell(i, 1);
                cell.SetValue(XLCellValue.FromObject(null));

                if (!loopInfo.LoopList.Any())
                {
                    if (loopInfo.IsInsertMode)
                    {
                        for (int j = 0; j < loopInfo.RowCopyCount; j++)
                        {
                            sheet.Row(i + j).Delete();
                        }
                    }
                    else 
                    {
                        //$Empty cells of strings beginning with $
                        for (int j = 1; j <= colCount; j++)
                        {
                            var x = sheet.Cell(i, j);
                            if (x.GetString().Trim().StartsWith("$"))
                            {
                                x.SetValue(XLCellValue.FromObject(null));
                            }
                        }
                        i++;
                    }
                    continue;
                }

                // copy rows
                CopyRows(sheet, i, loopInfo.RowCopyCount, loopInfo.LoopList.Count, loopInfo.IsInsertMode);

                // over write
                bool isFirstLoop = true;
                foreach (var e in loopInfo.LoopList)
                {
                    var elementConverter = converter.CreateChildExcelSymbolConverter(e, loopInfo.LoopName);

                    // Recursive Processing
                    var processedRows = await OverWrite(sheet, i, i + loopInfo.RowCopyCount - 1, colCount, elementConverter, pageLoopRowsInfoList);
                    i += processedRows;

                    // Increment endRow
                    endRow = IncrementEndRow(ref isFirstLoop, endRow, processedRows, loopInfo.RowCopyCount);
                }
            }
            // Processed Rows
            return endRow - startRow + 1;
        }

        class LoopInfo
        {
            internal int RowCopyCount { get; set; }
            internal List<object?> LoopList { get; set; } = new();
            internal string LoopName { get; set; } = string.Empty;
            internal bool IsInsertMode { get; set; }
        }

        static async Task<bool> TryParseLoop(string leftCell, IExcelSymbolConverter converter, LoopInfo loopInfo, string sheetName, List<PageLoopRowsInfo> pageLoopRowsInfoList)
        { 
            if (await TryParseLoopNormal(leftCell, converter, loopInfo)) return true;
            return TryParsePageLoop(leftCell, converter, loopInfo, sheetName, pageLoopRowsInfoList);
        }

        static async Task<bool> TryParseLoopNormal(string leftCell, IExcelSymbolConverter converter, LoopInfo loopInfo)
        {
            if (!leftCell.StartsWith("#LoopRow")) return false;

            // #LoopRow($list, i, rowCopyCount)
            var args = leftCell.Replace("#LoopRow", "").Replace("(", "").Replace(")", "").Split(',').Select(e => e.Trim()).ToArray();

            // rowCopyCount is optional
            var rowCopyCount = 1;
            if (args.Length == 3)
            {
                if (!int.TryParse(args[2], out rowCopyCount)) return false;
            }
            loopInfo.IsInsertMode = true;
            loopInfo.RowCopyCount = rowCopyCount;

            // #list and i(enumerable name) are must
            if (args.Length < 2) return false;

            if (!args[0].StartsWith("$")) return false;
            var enumerableName = args[0].Substring(1);
            loopInfo.LoopName = args[1];

            var enumerable = (await converter.GetData(enumerableName))?.Value as IEnumerable;
            if (enumerable == null) return false;

            loopInfo.LoopList = enumerable.OfType<object?>().ToList();

            return true;
        }

        static bool TryParsePageLoop(string leftCell, IExcelSymbolConverter converter, LoopInfo loopInfo, string sheetName, List<PageLoopRowsInfo> pageLoopRowsInfoList)
        {
            if (!leftCell.StartsWith("#PagedLoopRows")) return false;
            var args = leftCell.Replace("#PagedLoopRows", "").Replace("(", "").Replace(")", "").Split(',').Select(e => e.Trim()).ToArray();
            if (!int.TryParse(args[4], out var blockRowCount)) return false;
            loopInfo.IsInsertMode = false;

            var first = pageLoopRowsInfoList.FirstOrDefault(e => e.FirstPageSheetName == sheetName);
            if (first != null)
            {
                loopInfo.LoopList = first.FirstPageList;
                loopInfo.LoopName = args[3];
                loopInfo.RowCopyCount = blockRowCount;
                return true;
            }
            var body = pageLoopRowsInfoList.FirstOrDefault(e => e.BodyPageSheetNames.Contains(sheetName));
            if (body != null)
            {
                var index = body.BodyPageSheetNames.IndexOf(sheetName);
                loopInfo.LoopList = body.BodyPageLists[index];
                loopInfo.LoopName = args[3];
                loopInfo.RowCopyCount = blockRowCount;
                return true;
            }
            var last = pageLoopRowsInfoList.FirstOrDefault(e => e.LastPageSheetName == sheetName);
            if (last != null)
            {
                loopInfo.LoopList = last.LastPageList;
                loopInfo.LoopName = args[3];
                loopInfo.RowCopyCount = blockRowCount;
                return true;
            }
            return false;
        }

        static async Task OverWriteCell(IXLWorksheet sheet, int rowIndex, int colCount, Func<string, Task<ExcelOverWriteCell?>> converter)
        {
            for (var i = 0; i < colCount; i++)
            {
                var cellIndex = i + 1;
                var text = sheet.GetText(rowIndex, cellIndex).Trim();
                if (text.StartsWith("$"))
                {
                    var x = await converter(text.Substring(1));
                    SetCellData(sheet, rowIndex, cellIndex, x);
                }
            }
        }

        static void SetCellData(IXLWorksheet sheet, int rowIndex, int cellIndex, ExcelOverWriteCell? cellData)
        { 
            if (cellData == null) return;
            var cell = sheet.Cell(rowIndex, cellIndex);
            cell.SetValue(XLCellValue.FromObject(cellData.Value));
        }

        static void CopyRows(IXLWorksheet sheet, int rowIndex, int rowCopyCount, int loopCount, bool isInsertMode)
        {
            var rangeToCopy = sheet.Range($"{rowIndex}:{rowIndex + rowCopyCount - 1}");

            var srcHeights = new List<double>();
            for (int i = 0; i < rowCopyCount; i++)
            {
                srcHeights.Add(sheet.Row(rowIndex + i).Height);
            }
            for (int i = 1; i < loopCount; i++)
            {
                var insertRowIndex = rowIndex + rowCopyCount * i;
                var insertRow = isInsertMode ?
                    sheet.Row(insertRowIndex).InsertRowsAbove(rowCopyCount).First() :
                    sheet.Row(insertRowIndex);
                rangeToCopy.CopyTo(insertRow);
                for (int j = 0; j < rowCopyCount; j++)
                {
                    sheet.Row(insertRowIndex + j).Height = srcHeights[j];
                }
            }
        }

        static int IncrementEndRow(ref bool isFirstLoop, int endRow, int processedRows, int rowCopyCount)
        {
            if (isFirstLoop)
            {
                isFirstLoop = false;

                // Subtract duplicate rows from the processed rows
                endRow += (processedRows - rowCopyCount);
            }
            else
            {
                endRow += processedRows;
            }

            return endRow;
        }

    }
}
