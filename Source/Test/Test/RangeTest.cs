using ClosedXML.Excel;
using Excel.Report.PDF;

namespace Test
{
    public class RangeTest
    {
        [Test]
        public void Test()
        {
            //TODO excel

            using (var stream = new FileStream(Path.Combine(TestEnvironment.PdfSrcPath, "xxx.xlsx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var book = new XLWorkbook(stream))
            {
                //TODO Add rows

                book.SaveAs(Path.Combine(TestEnvironment.TestResultsPath, "xxx2.xlsx"));

            }

            using (var stream = new FileStream(Path.Combine(TestEnvironment.TestResultsPath, "xxx.xlsx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var book = new OpenClosedXML(stream))
            {
                book.GetSheetMaxRowCol(1, out var rowCount, out var colCount);

                //TODO check
            }
        }
    }
}
