using ClosedXML.Excel;

namespace Excel.Report.PDF
{
    public static class ExcelUtils
    {
        public static List<List<string>> ReadAllTexts(this IXLWorksheet sheet)
            => sheet.ReadAll((row, col) => sheet.GetText(row, col).Trim(), string.Empty);

        public static async Task<List<List<string>>> ReadAllTextsFromExcelBinary(Stream excel)
        {
            List<List<string>> texts = new();
            using (var memoryStream = new MemoryStream())
            {
                await excel.CopyToAsync(memoryStream);
                memoryStream.Position = 0;
                using (var book = new XLWorkbook(memoryStream))
                {
                    texts = book.Worksheet(1).ReadAllTexts();
                }
            }
            return texts;
        }

        public static List<List<object?>> ReadAllObjects(this IXLWorksheet sheet)
                 => sheet.ReadAll((row, col) => sheet.GetValue(row, col), null);

        public static MemoryStream CreateExcelBinary(List<List<string>> allTexts, string sheetName)
            => CreateExcelBinary(allTexts, sheetName, (cell, value) => { cell.SetValue(value).Style.NumberFormat.Format = "@"; });

        public static MemoryStream CreateExcelBinary(List<List<object?>> objects, string sheetName)
            => CreateExcelBinary(objects, sheetName, (cell, value) => cell.SetValue(XLCellValue.FromObject(value)));

        internal static string GetText(this IXLWorksheet sheet, int rowIndex, int columnIndex)
        {
            var cell = sheet.Cell(rowIndex, columnIndex);
            if (cell == null) return string.Empty;
            return cell.GetString();
        }

        internal static void GetRowColCount(this IXLWorksheet sheet, out int rowCount, out int colCount)
        {
            var usedRows = sheet.RowsUsed();
            rowCount = usedRows.OfType<IXLRow>().LastOrDefault()?.RowNumber() ?? 0;
            colCount = usedRows.OfType<IXLRow>().Select(e => e.CellsUsed().OfType<IXLCell>().LastOrDefault()?.Address?.ColumnNumber ?? 0).Max();
        }

        static List<List<T>> ReadAll<T>(this IXLWorksheet sheet, Func<int, int, T> getter, T defaultValue)
        {
            var rows = new List<List<T>>();
            GetRowColCount(sheet, out var rowCount, out var colCount);
            for (var i = 0; i < rowCount; i++)
            {
                var cols = new List<T>();
                var rowIndex = i + 1;
                for (var j = 0; j < colCount; j++)
                {
                    cols.Add(getter(rowIndex, j + 1));
                }
                cols.AddRange(Enumerable.Range(0, colCount - cols.Count).Select(_ => defaultValue));
                rows.Add(cols);
            }
            return rows;
        }

        static MemoryStream CreateExcelBinary<T>(List<List<T>> values, string sheetName, Action<IXLCell, T> setter)
        {
            using (var book = new XLWorkbook())
            {
                var sheet = book.AddWorksheet(sheetName);
                for (var i = 0; i < values.Count; i++)
                {
                    var rowIndex = i + 1;
                    for (var j = 0; j < values[i].Count; j++)
                    {
                        var cell = sheet.Cell(rowIndex, j + 1);
                        setter(cell, values[i][j]);
                    }
                }
                var stream = new MemoryStream();
                book.SaveAs(stream);
                stream.Seek(0, SeekOrigin.Begin);
                return stream;
            }
        }

        static object? GetValue(this IXLWorksheet sheet, int rowIndex, int columnIndex)
        {
            var cell = sheet.Cell(rowIndex, columnIndex);
            if (cell == null) return null;
            if (cell.IsEmpty()) return null;
            switch (cell.DataType)
            {
                case XLDataType.Text:
                    return cell.GetString();
                case XLDataType.Boolean:
                    return cell.GetBoolean();
                case XLDataType.DateTime:
                    return cell.GetDateTime();
                case XLDataType.Number:
                    return cell.GetDouble();
                case XLDataType.TimeSpan:
                    return cell.GetTimeSpan();
                default:
                    return null;
            }
        }
    }
}
