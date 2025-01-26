using ClosedXML.Excel;
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
        public void PageBreakRowColTest()
        {
            using (var stream = new FileStream(Path.Combine(TestEnvironment.PdfSrcPath, "PageBreakTest.xlsx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var book = new OpenClosedXML(stream))
            {
                var pageBreakInfo = PageBreakInfo.CreateRowColumnPageBreak(15, 5);
                var pages = book.GetCellInfo(1, 200, 200, out double scaling, pageBreakInfo);
                pages.Count.Is(6);

                {
                    var page = pages[0];
                    var pageLastCellInfo = page.Last();
                    pageLastCellInfo.Cell!.Address.RowNumber.Is(15);
                    pageLastCellInfo.Cell!.Address.ColumnNumber.Is(5);
                }
                {
                    var page = pages[1];
                    var pageLastCellInfo = page.Last();
                    pageLastCellInfo.Cell!.Address.RowNumber.Is(15);
                    pageLastCellInfo.Cell!.Address.ColumnNumber.Is(9);
                }
                {
                    var page = pages[2];
                    var pageLastCellInfo = page.Last();
                    pageLastCellInfo.Cell!.Address.RowNumber.Is(30);
                    pageLastCellInfo.Cell!.Address.ColumnNumber.Is(5);
                }
                {
                    var page = pages[3];
                    var pageLastCellInfo = page.Last();
                    pageLastCellInfo.Cell!.Address.RowNumber.Is(30);
                    pageLastCellInfo.Cell!.Address.ColumnNumber.Is(9);
                }
                {
                    var page = pages[4];
                    var pageLastCellInfo = page.Last();
                    pageLastCellInfo.Cell!.Address.RowNumber.Is(31);
                    pageLastCellInfo.Cell!.Address.ColumnNumber.Is(5);
                }
                {
                    var page = pages[5];
                    var pageLastCellInfo = page.Last();
                    pageLastCellInfo.Cell!.Address.RowNumber.Is(31);
                    pageLastCellInfo.Cell!.Address.ColumnNumber.Is(9);
                }

            }
        }

        [Test]
        public void PageBreakHighWidthTest()
        {
            using (var stream = new FileStream(Path.Combine(TestEnvironment.PdfSrcPath, "PageBreakTest.xlsx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var book = new OpenClosedXML(stream))
            {
                var pageBreakInfo = PageBreakInfo.CreateSizePageBreak(200, 35);
                var pages = book.GetCellInfo(1, 200, 200, out double scaling, pageBreakInfo);
                pages.Count.Is(6);

                {
                    var page = pages[0];
                    var pageLastCellInfo = page.Last();
                    pageLastCellInfo.Cell!.Address.RowNumber.Is(15);
                    pageLastCellInfo.Cell!.Address.ColumnNumber.Is(5);
                }
                {
                    var page = pages[1];
                    var pageLastCellInfo = page.Last();
                    pageLastCellInfo.Cell!.Address.RowNumber.Is(15);
                    pageLastCellInfo.Cell!.Address.ColumnNumber.Is(9);
                }
                {
                    var page = pages[2];
                    var pageLastCellInfo = page.Last();
                    pageLastCellInfo.Cell!.Address.RowNumber.Is(30);
                    pageLastCellInfo.Cell!.Address.ColumnNumber.Is(5);
                }
                {
                    var page = pages[3];
                    var pageLastCellInfo = page.Last();
                    pageLastCellInfo.Cell!.Address.RowNumber.Is(30);
                    pageLastCellInfo.Cell!.Address.ColumnNumber.Is(9);
                }
                {
                    var page = pages[4];
                    var pageLastCellInfo = page.Last();
                    pageLastCellInfo.Cell!.Address.RowNumber.Is(31);
                    pageLastCellInfo.Cell!.Address.ColumnNumber.Is(5);
                }
                {
                    var page = pages[5];
                    var pageLastCellInfo = page.Last();
                    pageLastCellInfo.Cell!.Address.RowNumber.Is(31);
                    pageLastCellInfo.Cell!.Address.ColumnNumber.Is(9);
                }

            }

        }
    }
}
