using ClosedXML.Excel;
using System;
using System.Collections;

namespace Excel.Report.PDF
{
    public static class ExcelOverWriter
    {
        public static async Task OverWrite(this IXLWorksheet sheet, IExcelSymbolConverter converter)
        {
            await Task.CompletedTask;

            // Get all rows and columns of the sheet
            ExcelUtils.GetRowColCount(sheet, out var rowCount, out var colCount);

            await OverWrite(sheet, 1, rowCount, colCount, converter);
        }

        private static async Task<int> OverWrite(IXLWorksheet sheet, int startRow, int endRow, int colCount, IExcelSymbolConverter converter)
        {
            for (int i = startRow; i <= endRow;)
            {
                await OverWriteCell(sheet, i, colCount, async t => await converter.GetData(t));

                LoopInfo loopInfo = new();
                if (!await TryParseLoop(sheet.GetText(i, 1).Trim(), converter, loopInfo))
                {
                    i++;
                    continue;
                }

                // delete #LoopRow
                var cell = sheet.Cell(i, 1);
                cell.SetValue(XLCellValue.FromObject(null));

                // copy rows
                CopyRows(sheet, i, loopInfo.RowCopyCount, loopInfo.LoopList.Count);

                // over write
                bool isFirstLoop = true;
                foreach (var e in loopInfo.LoopList)
                {
                    var elementConverter = converter.CreateChildExcelSymbolConverter(e, loopInfo.LoopName);

                    // Recursive Processing
                    var processedRows = await OverWrite(sheet, i, i + loopInfo.RowCopyCount - 1, colCount, elementConverter);
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
        }

        private static async Task<bool> TryParseLoop(string leftCell, IExcelSymbolConverter converter, LoopInfo loopInfo)
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

        static void CopyRows(IXLWorksheet sheet, int rowIndex, int rowCopyCount, int loopCount)
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
                var insertRow = sheet.Row(insertRowIndex).InsertRowsAbove(rowCopyCount).First();
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
