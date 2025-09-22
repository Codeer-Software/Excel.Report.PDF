using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Excel.Report.PDF;
using NUnit.Framework;
using PdfSharp.Fonts;

namespace Test
{
    public class ExcelOverWriterTest
    {
        class QuotationDetail
        {
            public string Title { get; set; } = string.Empty;
            public string Detail { get; set; } = string.Empty;
            public decimal Price { get; set; }
            public decimal Discount { get; set; }
            public decimal Total=>Price - Discount;
        }

        class Quotation
        {
            public string Title { get; set; } = string.Empty;
            public string Client { get; set; } = string.Empty;
            public string PersonInCharge { get; set; } = string.Empty;
            public List<QuotationDetail> Details { get; } = new();
            public decimal Total => Details.Sum(x => x.Total);
            public decimal Tax => Total * (decimal)0.1;
            public decimal TotalInTax => Total + Tax;
        }

        class Data
        {
            public List<Loop1> Loop1 { get; set; } = new();
            public string Name {  get; set; } = string.Empty;
        }

        class Loop1
        {
            public string Text { get; set; } = string.Empty;
            public List<Loop2> Loop2 { get; set; } = new();
        }

        class Loop2
        {
            public int Id { get; set; }
        }

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
        public async Task Test1()
        {
            var data = new Quotation 
            {
                Title = "宴会時の食材",
                Client = "エクセルコンサルティング株式会社",
                PersonInCharge = "大谷正一"
            };
            data.Details.Add(new()
            {
                Title = "鯛",
                Detail = "新鮮",
                Price = 10000,
                Discount = 0,
            });
            data.Details.Add(new()
            {
                Title = "鰤",
                Detail = "新鮮",
                Price = 20000,
                Discount = 0,
            });
            data.Details.Add(new()
            {
                Title = "ハマチ",
                Detail = "ご奉仕品",
                Price = 30000,
                Discount = 2000,
            });
            data.Details.Add(new()
            {
                Title = "蛸",
                Detail = "ご奉仕品",
                Price = 40000,
                Discount = 1000,
            });
            using (var stream = new FileStream(Path.Combine(TestEnvironment.PdfSrcPath, "Quotation.xlsx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var book = new XLWorkbook(stream))
            {
                await book.Worksheet(1).OverWrite(new ObjectExcelSymbolConverter(data));
                book.SaveAs(Path.Combine(TestEnvironment.TestResultsPath, "QuotationDst.xlsx"));

                var sheet = book.Worksheets.First();

                // B4:The part unrelated to the loop, The point where data is initially stored
                var noLoopFirstData = sheet.Cell(4, 2).Value.GetText();
                noLoopFirstData.Is("エクセルコンサルティング株式会社");

                // B18:The first line of the loop, Verify if the data is output as it is.
                var firstLoopData = sheet.Cell(18, 2).Value.GetText();
                firstLoopData.Is("鯛");

                // V21:The last loop, Check if the total value is stored
                var lastLoopSubtractData = sheet.Cell(21, 22).Value.GetNumber().ToString();
                lastLoopSubtractData.Is("39000");

                // V26:The last line, Check if the sum of each row and the tax is stored
                var lastData = sheet.Cell(26, 22).Value.GetNumber().ToString();
                lastData.Is("106700");


                // R14:Merging cells, Check if it is the same as the value in V26
                var total = sheet.Cell(14, 18).Value.GetNumber().ToString();
                total.Is("106700");

            }

            using var outStream = ExcelConverter.ConvertToPdf(Path.Combine(TestEnvironment.TestResultsPath, "QuotationDst.xlsx"), 1);
            File.WriteAllBytes(Path.Combine(TestEnvironment.TestResultsPath, "Quotation.pdf"), outStream.ToArray());
        }

        [Test]
        public async Task RecursiveNoLoopTest()
        {
            var data = new Data
            {
                Name ="NameA"
            };

            var loop1_1 = new Loop1
            {
                Text = "Test1"
            };

            loop1_1.Loop2.Add(new Loop2 { Id = 1 });
            loop1_1.Loop2.Add(new Loop2 { Id = 2 });
            loop1_1.Loop2.Add(new Loop2 { Id = 3 });

            var loop1_2 = new Loop1
            {
                Text = "Test2"
            };

            loop1_2.Loop2.Add(new Loop2 { Id = 11 });
            loop1_2.Loop2.Add(new Loop2 { Id = 22 });
            loop1_2.Loop2.Add(new Loop2 { Id = 33 });

            data.Loop1.Add(loop1_1);
            data.Loop1.Add(loop1_2);

            using (var stream = new FileStream(Path.Combine(TestEnvironment.PdfSrcPath, "RecursiveLoopTest(NoLoop).xlsx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var book = new XLWorkbook(stream))
            {
                await book.Worksheet(1).OverWrite(new ObjectExcelSymbolConverter(data));
                book.SaveAs(Path.Combine(TestEnvironment.TestResultsPath, "RecursiveLoopTest(NoLoop).xlsx"));

                var sheet = book.Worksheets.First();

                // B1:the part unrelated to the loop, verify if the data is output as it is.
                var noLoopData = sheet.Cell(1, 2).Value.GetText();
                noLoopData.Is("NameA");
            }

            using var outStream = ExcelConverter.ConvertToPdf(Path.Combine(TestEnvironment.TestResultsPath, "RecursiveLoopTest(NoLoop).xlsx"), 1);
            File.WriteAllBytes(Path.Combine(TestEnvironment.TestResultsPath, "RecursiveLoopTest(NoLoop).pdf"), outStream.ToArray());
        }

        [Test]
        public async Task RecursiveLoop1Test()
        {
            var data = new Data
            {
                Name = "NameA"
            };

            var loop1_1 = new Loop1
            {
                Text = "Test1"
            };

            loop1_1.Loop2.Add(new Loop2 { Id = 1 });
            loop1_1.Loop2.Add(new Loop2 { Id = 2 });
            loop1_1.Loop2.Add(new Loop2 { Id = 3 });

            var loop1_2 = new Loop1
            {
                Text = "Test2"
            };

            loop1_2.Loop2.Add(new Loop2 { Id = 11 });
            loop1_2.Loop2.Add(new Loop2 { Id = 22 });
            loop1_2.Loop2.Add(new Loop2 { Id = 33 });

            data.Loop1.Add(loop1_1);
            data.Loop1.Add(loop1_2);

            using (var stream = new FileStream(Path.Combine(TestEnvironment.PdfSrcPath, "RecursiveLoopTest(1Loop).xlsx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var book = new XLWorkbook(stream))
            {
                await book.Worksheet(1).OverWrite(new ObjectExcelSymbolConverter(data));
                book.SaveAs(Path.Combine(TestEnvironment.TestResultsPath, "RecursiveLoopTest(1Loop).xlsx"));

                var sheet = book.Worksheets.First();

                // B1:the part unrelated to the loop, verify if the data is output as it is.
                var noLoopData = sheet.Cell(1, 2).Value.GetText();
                noLoopData.Is("NameA");

                // B2:The first line of the loop, verify if the data is output as it is.
                var firstLoopData = sheet.Cell(2, 2).Value.GetText();
                firstLoopData.Is("Test1");

                // B3:The last line of the loop, verify if the data is output as it is.
                var lastLoopData = sheet.Cell(3, 2).Value.GetText();
                lastLoopData.Is("Test2");
            }

            using var outStream = ExcelConverter.ConvertToPdf(Path.Combine(TestEnvironment.TestResultsPath, "RecursiveLoopTest(1Loop).xlsx"), 1);
            File.WriteAllBytes(Path.Combine(TestEnvironment.TestResultsPath, "RecursiveLoopTest(1Loop).pdf"), outStream.ToArray());
        }

        [Test]
        public async Task RecursiveLoop2Test()
        {
            var data = new Data
            {
                Name = "NameA"
            };

            var loop1_1 = new Loop1
            {
                Text = "Test1"
            };

            loop1_1.Loop2.Add(new Loop2 { Id = 1 });
            loop1_1.Loop2.Add(new Loop2 { Id = 2 });
            loop1_1.Loop2.Add(new Loop2 { Id = 3 });

            var loop1_2 = new Loop1
            {
                Text = "Test2"
            };

            loop1_2.Loop2.Add(new Loop2 { Id = 11 });
            loop1_2.Loop2.Add(new Loop2 { Id = 22 });
            loop1_2.Loop2.Add(new Loop2 { Id = 33 });

            data.Loop1.Add(loop1_1);
            data.Loop1.Add(loop1_2);

            using (var stream = new FileStream(Path.Combine(TestEnvironment.PdfSrcPath, "RecursiveLoopTest(2Loop).xlsx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var book = new XLWorkbook(stream))
            {
                await book.Worksheet(1).OverWrite(new ObjectExcelSymbolConverter(data));
                book.SaveAs(Path.Combine(TestEnvironment.TestResultsPath, "RecursiveLoopTest(2Loop).xlsx"));

                var sheet = book.Worksheets.First();

                // B1:the part unrelated to the loop, verify if the data is output as it is.
                var noLoopData = sheet.Cell(1, 2).Value.GetText();
                noLoopData.Is("NameA");

                // B2:The first iteration of Loop1, verify if the value is stored as it is.
                var firstIterationLoop1Data = sheet.Cell(2, 2).Value.GetText();
                firstIterationLoop1Data.Is("Test1");

                // B3:The first iteration of Loop2 within the first iteration of Loop1, verify if the value is stored as it is.
                var firstLoop2DatawithinFirstLoop1 = sheet.Cell(3, 2).Value.GetNumber().ToString();
                firstLoop2DatawithinFirstLoop1.Is("1");

                // B5:The last iteration of Loop2 within the first iteration of Loop1, verify if the value is stored as it is.
                var lastLoop2DatawithinFirstLoop1 = sheet.Cell(5, 2).Value.GetNumber().ToString();
                lastLoop2DatawithinFirstLoop1.Is("3");

                // B6:The last iteration of Loop1, verify if the value is stored as it is.
                var lastIterationLoop1Data = sheet.Cell(6, 2).Value.GetText();
                lastIterationLoop1Data.Is("Test2");

                // B7:The first iteration of Loop2 within the last iteration of Loop1, verify if the value is stored as it is.
                var firstLoop2DatawithinLastLoop1 = sheet.Cell(7, 2).Value.GetNumber().ToString();
                firstLoop2DatawithinLastLoop1.Is("11");

                // B9:The last iteration of Loop2 within the last iteration of Loop1, verify if the value is stored as it is.
                var lastLoop2DatawithinLastLoop1 = sheet.Cell(9, 2).Value.GetNumber().ToString();
                lastLoop2DatawithinLastLoop1.Is("33");
            }

            using var outStream = ExcelConverter.ConvertToPdf(Path.Combine(TestEnvironment.TestResultsPath, "RecursiveLoopTest(2Loop).xlsx"), 1);
            File.WriteAllBytes(Path.Combine(TestEnvironment.TestResultsPath, "RecursiveLoopTest(2Loop).pdf"), outStream.ToArray());
        }

        [Test]
        public void TestCopyPage()
        {
            using (var stream = new FileStream(Path.Combine(TestEnvironment.PdfSrcPath, "TestCopyPage.xlsx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var book = new XLWorkbook(stream))
            {   
                var firstSheet = book.Worksheet(1);
                var src = firstSheet.Name;
                firstSheet.Name = src + "_" + 1;
                for (int i = 1; i <= 3; i++)
                {
                    var copy = firstSheet.CopyTo($"{src}_{i + 1}");
                }
                book.SaveAs(Path.Combine(TestEnvironment.TestResultsPath, "TestCopyPage.xlsx"));
            }

            using var outStream = ExcelConverter.ConvertToPdf(Path.Combine(TestEnvironment.TestResultsPath, "TestCopyPage.xlsx"));
            File.WriteAllBytes(Path.Combine(TestEnvironment.TestResultsPath, "TestCopyPage.pdf"), outStream.ToArray());

        }

        class SimpleDataOwner
        { 
            public List<SimpleData> Details { get; set; } = new();
        }

        class SimpleData
        { 
            public string Text { get; set; } = string.Empty;
            public int Number { get; set; }
        }

        [Test]
        public async Task MultiPageSheetTest1()
        {
            var data = new SimpleDataOwner();

            for (int i = 0; i < 100; i++)
            {
                data.Details.Add(new SimpleData { Text = $"Test{i + 1}", Number = i + 1 });
            }

            using (var stream = new FileStream(Path.Combine(TestEnvironment.PdfSrcPath, "MultiPageSheetTest.xlsx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var book = new XLWorkbook(stream))
            {
                await book.OverWrite(new ObjectExcelSymbolConverter(data));
                book.SaveAs(Path.Combine(TestEnvironment.TestResultsPath, "MultiPageSheetTest1.xlsx"));

                book.Worksheets.Count.Is(5);
                book.Worksheet("first").Cell(11, 2).Value.Is("Test10");
                book.Worksheet("body_0").Cell(31, 2).Value.Is("Test40");
                book.Worksheet("last").Cell(2, 2).Value.Is("★");
            }

            using var outStream = ExcelConverter.ConvertToPdf(Path.Combine(TestEnvironment.TestResultsPath, "MultiPageSheetTest1.xlsx"));
            File.WriteAllBytes(Path.Combine(TestEnvironment.TestResultsPath, "MultiPageSheetTest1.pdf"), outStream.ToArray());
        }

        [Test]
        public async Task MultiPageSheetTest2()
        {
            var data = new SimpleDataOwner();

            for (int i = 0; i < 110; i++)
            {
                data.Details.Add(new SimpleData { Text = $"Test{i + 1}", Number = i + 1 });
            }

            using (var stream = new FileStream(Path.Combine(TestEnvironment.PdfSrcPath, "MultiPageSheetTest.xlsx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var book = new XLWorkbook(stream))
            {
                await book.OverWrite(new ObjectExcelSymbolConverter(data));
                book.SaveAs(Path.Combine(TestEnvironment.TestResultsPath, "MultiPageSheetTest2.xlsx"));

                book.Worksheets.Count.Is(5);
                book.Worksheet("first").Cell(11, 2).Value.Is("Test10");
                book.Worksheet("body_0").Cell(31, 2).Value.Is("Test40");
                book.Worksheet("last").Cell(11, 2).Value.Is("Test110");
                book.Worksheet("last").Cell(12, 2).Value.Is("★");
            }

            using var outStream = ExcelConverter.ConvertToPdf(Path.Combine(TestEnvironment.TestResultsPath, "MultiPageSheetTest2.xlsx"));
            File.WriteAllBytes(Path.Combine(TestEnvironment.TestResultsPath, "MultiPageSheetTest2.pdf"), outStream.ToArray());
        }

        [Test]
        public async Task MultiPageSheetBodyLastTest1()
        {
            var data = new SimpleDataOwner();

            for (int i = 0; i < 90; i++)
            {
                data.Details.Add(new SimpleData { Text = $"Test{i + 1}", Number = i + 1 });
            }

            using (var stream = new FileStream(Path.Combine(TestEnvironment.PdfSrcPath, "MultiPageSheetBodyLastTest.xlsx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var book = new XLWorkbook(stream))
            {
                await book.OverWrite(new ObjectExcelSymbolConverter(data));
                book.SaveAs(Path.Combine(TestEnvironment.TestResultsPath, "MultiPageSheetBodyLastTest1.xlsx"));

                book.Worksheets.Count.Is(4);
                book.Worksheet("body_0").Cell(2, 2).Value.Is("Test1");
                book.Worksheet("body_0").Cell(31, 2).Value.Is("Test30");
                book.Worksheet("last").Cell(2, 2).Value.Is("★");
            }

            using var outStream = ExcelConverter.ConvertToPdf(Path.Combine(TestEnvironment.TestResultsPath, "MultiPageSheetBodyLastTest1.xlsx"));
            File.WriteAllBytes(Path.Combine(TestEnvironment.TestResultsPath, "MultiPageSheetBodyLastTest1.pdf"), outStream.ToArray());
        }

        [Test]
        public async Task MultiPageSheetBodyLastTest2()
        {
            var data = new SimpleDataOwner();

            for (int i = 0; i < 100; i++)
            {
                data.Details.Add(new SimpleData { Text = $"Test{i + 1}", Number = i + 1 });
            }

            using (var stream = new FileStream(Path.Combine(TestEnvironment.PdfSrcPath, "MultiPageSheetBodyLastTest.xlsx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var book = new XLWorkbook(stream))
            {
                await book.OverWrite(new ObjectExcelSymbolConverter(data));
                book.SaveAs(Path.Combine(TestEnvironment.TestResultsPath, "MultiPageSheetBodyLastTest2.xlsx"));

                book.Worksheets.Count.Is(4);
                book.Worksheet("body_0").Cell(2, 2).Value.Is("Test1");
                book.Worksheet("body_0").Cell(31, 2).Value.Is("Test30");
                book.Worksheet("last").Cell(11, 2).Value.Is("Test100");
                book.Worksheet("last").Cell(12, 2).Value.Is("★");
            }

            using var outStream = ExcelConverter.ConvertToPdf(Path.Combine(TestEnvironment.TestResultsPath, "MultiPageSheetBodyLastTest2.xlsx"));
            File.WriteAllBytes(Path.Combine(TestEnvironment.TestResultsPath, "MultiPageSheetBodyLastTest2.pdf"), outStream.ToArray());
        }
    }
}
