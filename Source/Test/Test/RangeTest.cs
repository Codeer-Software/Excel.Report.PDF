using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Excel.Report.PDF;
using Microsoft.VisualStudio.TestPlatform.PlatformAbstractions.Interfaces;
using NUnit.Framework.Internal.Execution;
using PdfSharp.Fonts;
using System.IO;
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
            using (var book = new OpenClosedXML(stream))
            {
                var pageBreakInfo = new OpenClosedXML.PageBreakInfo(true, 15, 5);
                var pages = book.GetCellInfo(1, 200, 200, out double scaling, pageBreakInfo);
                pages.Count.Is(4);

                {
                    var page = pages[0];
                    var pageLastCellInfo = page.Last();
                    pageLastCellInfo.Cell!.Address.RowNumber.Is(15);
                }
                {
                    var page = pages[1];
                    var pageLastCellInfo = page.Last();
                    pageLastCellInfo.Cell!.Address.RowNumber.Is(30);
                }
                {
                    var page = pages[2];
                    var pageLastCellInfo = page.Last();
                    pageLastCellInfo.Cell!.Address.RowNumber.Is(30);
                }
                {
                    var page = pages[3];
                    var pageLastCellInfo = page.Last();
                    pageLastCellInfo.Cell!.Address.RowNumber.Is(30);
                }

                //    string? cellInfo2 = pages[0]?.LastOrDefault()?.Cell?.ToString();
                //  var pageBreakColum = cellInfo1?.Where(char.IsLetter).ToArray();
                //  pageBreakColum?.Is("E");

            }

           
        }
    }
}
