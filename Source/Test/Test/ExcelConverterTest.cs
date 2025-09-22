using Excel.Report.PDF;
using PdfSharp.Fonts;

namespace Test
{
    public class ExcelConverterTest
    {
        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            if (GlobalFontSettings.FontResolver == null) GlobalFontSettings.FontResolver = new CustomFontResolver();

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

        [Test]
        public void Test6() => Convert("Test6");

        [Test]
        public void Test7() => Convert("Test7");

        [Test]
        public void Test8() => Convert("Test8");

        [Test]
        public void Test9() => Convert("Test9");

        [Test]
        public void TestMargin() => Convert("TestMargin");

        [Test]
        public void Testmultipage() => Convert("multipagetest");

        [Test]
        public void TestDirectionMarginScaling() => Convert("TestDirectionMarginScaling");

        [Test]
        public void TestMarginScaling() => Convert("TestMarginScaling");

        [Test]
        public void TestDoubleLine() => Convert("TestDoubleLine");

        [Test]
        public void TestVirticalText() => Convert("TestVirticalText");

        [Test]
        public void TestBarCode() => Convert("TestBarCode");

        [Test]
        public void TestMultiSheet()
        {
            var workbookPath = Path.Combine(TestEnvironment.PdfSrcPath, "TestMultiSheet.xlsx");
            using var outStream = ExcelConverter.ConvertToPdf(workbookPath);
            File.WriteAllBytes(Path.Combine(TestEnvironment.TestResultsPath, "TestMultiSheet.pdf"), outStream.ToArray());
        }

        [Test]
        public void PageBreakRowColTest()
        {
            var pageBreakInfo = PageBreakInfo.CreateRowColumnPageBreak(15, 5);
            Convert("PageBreakTest", pageBreakInfo);
        }

        [Test]
        public void PageBreakHighWidthTest()
        {
            var pageBreakInfo = PageBreakInfo.CreateSizePageBreak(200, 35);
            Convert("PageBreakHighWidthTest", pageBreakInfo);
        }

        public void Convert(string name, PageBreakInfo? pageBreakInfo = null)
        {
            var workbookPath = Path.Combine(TestEnvironment.PdfSrcPath, $"{name}.xlsx");
            using var outStream = ExcelConverter.ConvertToPdf(workbookPath, 1, pageBreakInfo);
            File.WriteAllBytes(Path.Combine(TestEnvironment.TestResultsPath, $"{name}.pdf"), outStream.ToArray());
        }
    }
}
