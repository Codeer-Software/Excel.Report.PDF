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

        [Test]
        public void Test2() => Convert("Test2");

        [Test]
        public void Test3() => Convert("Test3");

        [Test]
        public void Test4() => Convert("Test4");

        [Test]
        public void Test5() => Convert("Test5");

        public void Convert(string name)
        {
            var workbookPath = Path.Combine(TestEnvironment.PdfSrcPath, $"{name}.xlsx");
            using var outStream = ExcelConverter.ConvertToPdf(workbookPath, 1);
            File.WriteAllBytes(Path.Combine(TestEnvironment.TestResultsPath, $"{name}.pdf"), outStream.ToArray());
        }
    }
}
