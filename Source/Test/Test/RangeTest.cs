using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Excel.Report.PDF;
using PdfSharp.Fonts;
namespace Test
{
    public class RangeTest
    {
        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            GlobalFontSettings.FontResolver = new CustomFontResolver();

            if (Directory.Exists(TestEnvironment.TestResultsPath))
            {
                Directory.Delete(TestEnvironment.TestResultsPath, true);
            }
            Directory.CreateDirectory(TestEnvironment.TestResultsPath);
        }

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

        [Test]
        public void PageBreakTest()
        {
            using (var stream = new FileStream(Path.Combine(TestEnvironment.PdfSrcPath, "PageBreakTest.xlsx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var book = new XLWorkbook(stream))
            {
                var worksheet = book.Worksheet(1);
                worksheet.Row(1).InsertRowsAbove(1);
                worksheet.Cell("A1").Value = "NewCellData";

                var openClosedXML = new OpenClosedXML(stream);
                // Pass page break information (row and column) as arguments
                openClosedXML.GetPageRanges(worksheet, 1, 10, 2);

                book.SaveAs(Path.Combine(TestEnvironment.TestResultsPath, "PageBreakTest.xlsx"));
            }

            using (var document = SpreadsheetDocument.Open(Path.Combine(TestEnvironment.TestResultsPath, "PageBreakTest.xlsx"), false))
            {
                var workbookPart = document.WorkbookPart;
                var sheet = workbookPart!.Workbook.Sheets!.Elements<Sheet>().FirstOrDefault();

                // Specify sheet1
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById("rId1");
                var worksheet = worksheetPart.Worksheet;

                // Get the page break setting
                var rowBreaks = worksheet.Elements<RowBreaks>().FirstOrDefault();
                var colBreaks = worksheet.Elements<ColumnBreaks>().FirstOrDefault();

                // Get the last element of the page break setting (row)
                var lastRowBreak = rowBreaks?.LastChild as Break;
                var lastRowBreakId = lastRowBreak?.Id ?? 0;

                // Get the last element of the page break setting (column)
                var lastColBreaks = colBreaks?.LastChild as Break;
                var lastColBreakId = lastColBreaks?.Id ?? 0;

                lastRowBreakId.InnerText.Is("10");
                lastColBreakId.InnerText.Is("2");

                using var outStream = ExcelConverter.ConvertToPdf(Path.Combine(TestEnvironment.TestResultsPath, "PageBreakTest.xlsx"), 1);
                File.WriteAllBytes(Path.Combine(TestEnvironment.TestResultsPath, "PageBreakTest.pdf"), outStream.ToArray());
            }
        }
    }
}
