using ClosedXML.Excel;
using Excel.Report.PDF;

namespace Test
{
    public class RangeTest
    {
        [Test]
        public void AddRowTest()
        {
            using (var stream = new FileStream(Path.Combine(TestEnvironment.PdfSrcPath, "RangeTestAddRow.xlsx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var book = new XLWorkbook(stream))
            {
                // Add rows
                var worksheet = book.Worksheet(1);
                worksheet.Row(1).InsertRowsAbove(1);
                worksheet.Cell("A1").Value = "NewCellData";

                book.SaveAs(Path.Combine(TestEnvironment.TestResultsPath, "ResultRangeTestAddRow.xlsx"));

            }

            using (var stream = new FileStream(Path.Combine(TestEnvironment.TestResultsPath, "ResultRangeTestAddRow.xlsx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var book = new OpenClosedXML(stream))
            {
                book.GetSheetMaxRowCol(1, out var rowCount, out var colCount);

                // Check whether the number of rows and columns can be obtained correctly even if a row is added
                rowCount.Is(5);
                colCount.Is(1);
            }
        }

        [Test]
        public void StyleCellTest()
        {
            using (var stream = new FileStream(Path.Combine(TestEnvironment.PdfSrcPath, "RangeTestStyleCell.xlsx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var book = new XLWorkbook(stream))
            {
                book.SaveAs(Path.Combine(TestEnvironment.TestResultsPath, "ResultRangeTestStyleCell.xlsx"));
            }

            using (var stream = new FileStream(Path.Combine(TestEnvironment.TestResultsPath, "ResultRangeTestStyleCell.xlsx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var book = new OpenClosedXML(stream))
            {
                book.GetSheetMaxRowCol(1, out var rowCount, out var colCount);

                // Check whether the number of rows and columns can be obtained correctly even if there is a cell with only style
                rowCount.Is(10);
                colCount.Is(1);
            }
        }
    }
}
