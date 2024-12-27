using ClosedXML.Excel;
using System.Collections;

namespace Excel.Report.PDF
{
    public static class ExcelOverWriter
    {
        private static async Task<int> OverWrite(IXLWorksheet sheet, int startRow, int endRow, int colCount, IExcelSymbolConverter converter)
        {
            for (int i = startRow; i <= endRow;)
            {

                await OverWriteCell(sheet, i, colCount, async t => await converter.GetData(t));

                // Get the string in column A of row i
                var leftCell = sheet.GetText(i, 1).Trim();
                if (!leftCell.StartsWith("#LoopRow"))
                {
                    i++;
                    continue;
                }

                // #LoopRow($list, i, rowCopyCount)
                var args = leftCell.Replace("#LoopRow", "").Replace("(", "").Replace(")", "").Split(',').Select(e => e.Trim()).ToArray();

                // rowCopyCount is optional
                var rowCopyCount = 1;
                if (args.Length == 3)
                {
                    if (!int.TryParse(args[2], out rowCopyCount)) continue;
                }

                // #list and i(enumerable name) are must
                if (args.Length < 2) continue;

                if (!args[0].StartsWith("$")) continue;

                var enumerable = (await converter.GetData(args[0].Substring(1)))?.Value as IEnumerable;
                if (enumerable == null) continue;

                var list = enumerable.OfType<object?>().ToList();

                // delete #LoopRow
                var cell = sheet.Cell(i, 1);
                cell.SetValue(XLCellValue.FromObject(null));

                // copy rows
                CopyRows(sheet, i, rowCopyCount, list.Count);

                // over write
                bool first = true;
                foreach (var e in list)
                {
                    var elementConverter = converter.CreateChildExcelSymbolConverter(e, args[1]);

                    // Recursive Processing
                    var processedRows = await OverWrite(sheet, i, i + rowCopyCount - 1, colCount, elementConverter);
                    i += processedRows;

                    if (first)
                    {
                        first = false;

                        // Subtract duplicate rows from the processed rows
                        endRow += (processedRows - rowCopyCount);
                    }
                    else
                    {
                        endRow += processedRows;
                    }
                }
            }
            // Processed Rows
            return endRow - startRow + 1;
        }

        public static async Task OverWrite(this IXLWorksheet sheet, IExcelSymbolConverter converter)
        {
            await Task.CompletedTask;

            // Get all rows and columns of the sheet
            ExcelUtils.GetRowColCount(sheet, out var rowCount, out var colCount);

            await OverWrite(sheet, 1, rowCount, colCount, converter);
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
    }
}
