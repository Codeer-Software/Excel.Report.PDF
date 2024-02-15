using Excel.Report.PDF;
using PdfSharp.Fonts;

namespace Test
{
    public class ExcelConverterTest
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
        public void Test1() => Convert("Test1");

        public void Convert(string name)
        {
            var workbookPath = Path.Combine(TestEnvironment.PdfSrcPath, $"{name}.xlsx");
            using var stream = new FileStream(workbookPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var outStream = ExcelConverter.ConvertToPdf(stream, 1);
            File.WriteAllBytes(Path.Combine(TestEnvironment.TestResultsPath, $"{name}.pdf"), outStream.ToArray());
        }
    }
}
