using ClosedXML.Excel;
using System.Collections;

namespace Excel.Report.PDF
{
    public static class ExcelOverWriter
    {
        public static async Task OverWrite(this IXLWorksheet sheet, IExcelSymbolConverter converter)
        {
            ExcelUtils.GetRowColCount(sheet, out var rowCount, out var colCount);
            for (var i = 0; i < rowCount; i++)
            {
                var rowIndex = i + 1;

                var leftCell = sheet.GetText(rowIndex, 1).Trim();

                if (leftCell.StartsWith("#LoopRow"))
                {
                    // #LoopRow($list, i, rowCopyCount)
                    var args = leftCell.Replace("#LoopRow", "").Replace("(", "").Replace(")", "").Split(',').Select(e=>e.Trim()).ToArray();

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
                    var enumerableName = args[1];

                    // delete #LoopRow
                    var cell = sheet.Cell(rowIndex, 1);
                    cell.SetValue(XLCellValue.FromObject(null));

                    // copy rows
                    CopyRows(sheet, rowIndex, rowCopyCount, list.Count);

                    // over write
                    foreach (var e in list)
                    {
                        for (int j = 0; j < rowCopyCount; j++)
                        {
                            await OverWriteCell(sheet, rowIndex, colCount, async t => await converter.GetData(e, enumerableName, t));
                            rowIndex++;
                        }
                    }

                    var addRowCount = (list.Count - 1) * rowCopyCount;
                    rowCount += addRowCount;
                    i += addRowCount;
                }
                else
                {
                    await OverWriteCell(sheet, rowIndex, colCount, async t => await converter.GetData(t));
                }
            }
        }

        static async Task OverWriteCell(IXLWorksheet sheet, int rowIndex, int colCount, Func<string, Task<ExcelOverWriteCell?>> converter)
        {
            for (var i = 0; i < colCount; i++)
            {
                var cellIndex = i + 1;
                var text = sheet.GetText(rowIndex, cellIndex).Trim();
                if (text.StartsWith("$"))
                {
                    SetCellData(sheet, rowIndex, cellIndex, await converter(text.Substring(1)));
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
