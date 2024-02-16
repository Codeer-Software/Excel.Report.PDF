using ClosedXML.Excel;
using Excel.Report.PDF;
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
            }

            using var outStream = ExcelConverter.ConvertToPdf(Path.Combine(TestEnvironment.TestResultsPath, "QuotationDst.xlsx"), 1);
            File.WriteAllBytes(Path.Combine(TestEnvironment.TestResultsPath, "Quotation.pdf"), outStream.ToArray());
        }

    }
}
