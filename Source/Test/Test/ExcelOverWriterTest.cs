using ClosedXML.Excel;
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
            GlobalFontSettings.FontResolver = new CustomFontResolver();

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

                // B4:Loopと関係ない部分、一番最初にデータが入るところ
                var noLoopFirstData = sheet.Cell(4, 2).Value.GetText();
                noLoopFirstData.Is("エクセルコンサルティング株式会社");

                // B18:Loopの1行目、データをそのまま出力
                var firstLoopData = sheet.Cell(18, 2).Value.GetText();
                firstLoopData.Is("鯛");

                // V21:Loopの最後、計算した値が入っているか確認
                var lastLoopSubtractData = sheet.Cell(21, 22).Value.GetNumber().ToString();
                lastLoopSubtractData.Is("39000");

                // V26:最後の行、合計が合っているか確認
                var lastData = sheet.Cell(26, 22).Value.GetNumber().ToString();
                lastData.Is("106700");


                // R14:セル結合を行っている、V26の値と同じか
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
                Text = "テキスト1"
            };

            loop1_1.Loop2.Add(new Loop2 { Id = 1 });
            loop1_1.Loop2.Add(new Loop2 { Id = 2 });
            loop1_1.Loop2.Add(new Loop2 { Id = 3 });

            var loop1_2 = new Loop1
            {
                Text = "テキスト2"
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
                Text = "テキスト1"
            };

            loop1_1.Loop2.Add(new Loop2 { Id = 1 });
            loop1_1.Loop2.Add(new Loop2 { Id = 2 });
            loop1_1.Loop2.Add(new Loop2 { Id = 3 });

            var loop1_2 = new Loop1
            {
                Text = "テキスト2"
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
                Text = "テキスト1"
            };

            loop1_1.Loop2.Add(new Loop2 { Id = 1 });
            loop1_1.Loop2.Add(new Loop2 { Id = 2 });
            loop1_1.Loop2.Add(new Loop2 { Id = 3 });

            var loop1_2 = new Loop1
            {
                Text = "テキスト2"
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
            }

            using var outStream = ExcelConverter.ConvertToPdf(Path.Combine(TestEnvironment.TestResultsPath, "RecursiveLoopTest(2Loop).xlsx"), 1);
            File.WriteAllBytes(Path.Combine(TestEnvironment.TestResultsPath, "RecursiveLoopTest(2Loop).pdf"), outStream.ToArray());

        }


    }
}
