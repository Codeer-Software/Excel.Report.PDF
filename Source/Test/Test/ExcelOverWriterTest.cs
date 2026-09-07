using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Excel.Report.PDF;
using NUnit.Framework;
using PdfSharp.Fonts;
using Test.Properties;

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

        class DataManyProps
        {
            public string Prop1 { get; set; } = "Prop1";
            public string Prop2 { get; set; } = "Prop2";
            public string Prop3 { get; set; } = "Prop3";
            public string Prop4 { get; set; } = "Prop4";
            public string Prop5 { get; set; } = "Prop5";
            public string Prop6 { get; set; } = "Prop6";
            public string Prop7 { get; set; } = "Prop7";
            public string Prop8 { get; set; } = "Prop8";
            public string Prop9 { get; set; } = "Prop9";
            public string Prop10 { get; set; } = "Prop10";

            public string Prop11 { get; set; } = "Prop11";
            public string Prop12 { get; set; } = "Prop12";
            public string Prop13 { get; set; } = "Prop13";
            public string Prop14 { get; set; } = "Prop14";
            public string Prop15 { get; set; } = "Prop15";
            public string Prop16 { get; set; } = "Prop16";
            public string Prop17 { get; set; } = "Prop17";
            public string Prop18 { get; set; } = "Prop18";
            public string Prop19 { get; set; } = "Prop19";
            public string Prop20 { get; set; } = "Prop20";

            public string Prop21 { get; set; } = "Prop21";
            public string Prop22 { get; set; } = "Prop22";
            public string Prop23 { get; set; } = "Prop23";
            public string Prop24 { get; set; } = "Prop24";
            public string Prop25 { get; set; } = "Prop25";
            public string Prop26 { get; set; } = "Prop26";
            public string Prop27 { get; set; } = "Prop27";
            public string Prop28 { get; set; } = "Prop28";
            public string Prop29 { get; set; } = "Prop29";
            public string Prop30 { get; set; } = "Prop30";

            public string Prop31 { get; set; } = "Prop31";
            public string Prop32 { get; set; } = "Prop32";
            public string Prop33 { get; set; } = "Prop33";
            public string Prop34 { get; set; } = "Prop34";
            public string Prop35 { get; set; } = "Prop35";
            public string Prop36 { get; set; } = "Prop36";
            public string Prop37 { get; set; } = "Prop37";
            public string Prop38 { get; set; } = "Prop38";
            public string Prop39 { get; set; } = "Prop39";
            public string Prop40 { get; set; } = "Prop40";
        }

        const string EmptySheetInputFileName = "EmptySheetTest.xlsx";
        const string RecursiveLoop2TestInputFileName = "ExcelOverWriterTest_RecursiveLoop2Test.xlsx";
        const string RecursiveLoop2TestMergedInputFileName = "ExcelOverWriterTest_RecursiveLoop2Test(Merged).xlsx";
        const string RecursiveLoop2TestManyRowsAndPropsInputFileName = "ExcelOverWriterTest_RecursiveLoop2Test(ManyRowsAndProps).xlsx";

        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            if (GlobalFontSettings.FontResolver == null) GlobalFontSettings.FontResolver = new CustomFontResolver();

            if (Directory.Exists(TestEnvironment.TestResultsPath))
            {
                Directory.Delete(TestEnvironment.TestResultsPath, true);
            }
            Directory.CreateDirectory(TestEnvironment.TestResultsPath);

            var emptySheetInputPath = Path.Combine(TestEnvironment.PdfSrcPath, EmptySheetInputFileName);
            if (!File.Exists(emptySheetInputPath))
            {
                Directory.CreateDirectory(TestEnvironment.PdfSrcPath);
                CreateEmptySheetInputWorkbook(emptySheetInputPath);
            }

            // Generate input files if they do not exist
            var recursiveLoop2TestInputPath = Path.Combine( TestEnvironment.PdfSrcPath, RecursiveLoop2TestInputFileName );
            if (!File.Exists( recursiveLoop2TestInputPath ))
            {
                Directory.CreateDirectory( TestEnvironment.PdfSrcPath );
                RecursiveLoop2TestInputWorkbook( recursiveLoop2TestInputPath );
            }

            var recursiveLoop2TestMergedInputPath = Path.Combine( TestEnvironment.PdfSrcPath, RecursiveLoop2TestMergedInputFileName );
            if (!File.Exists( recursiveLoop2TestMergedInputPath ))
            {
                Directory.CreateDirectory( TestEnvironment.PdfSrcPath );
                RecursiveLoop2TestMergedInputWorkbook( recursiveLoop2TestMergedInputPath );
            }

            var recursiveLoop2TestManyRowsAndPropsInputPath = Path.Combine( TestEnvironment.PdfSrcPath, RecursiveLoop2TestManyRowsAndPropsInputFileName );
            if (!File.Exists( recursiveLoop2TestManyRowsAndPropsInputPath )) {
                Directory.CreateDirectory( TestEnvironment.PdfSrcPath );
                RecursiveLoop2TestManyRowsAndPropsInputWorkbook( recursiveLoop2TestManyRowsAndPropsInputPath );
            }

        }

        private void RecursiveLoop2TestInputWorkbook(string path)
        {
            using var book = new XLWorkbook();
            var sheet = book.AddWorksheet( "Sheet1" );

            // Header row (literal — outside the loop).
            sheet.Cell( 1, 2 ).SetValue( "$Name" );

            // 2-row #LoopRow block (insert mode). Row 2 is the directive, row 3 is the
            // block's second row.
            sheet.Column( 1 ).Width = 30;
            sheet.Cell( 2, 1 ).SetValue( "#LoopRow($Loop1, x, 2)" );
            sheet.Cell( 2, 2 ).SetValue( "$x.Text" );
            sheet.Cell( 2, 4 ).Style.Fill.BackgroundColor = XLColor.Yellow;
            sheet.Cell( 3, 1 ).SetValue( "#LoopRow($x.Loop2, y, 1)" );
            sheet.Cell( 3, 2 ).SetValue( "$y.Id" );
            sheet.Cell( 3, 4 ).Style.Fill.BackgroundColor = XLColor.Red;

            // filter on C1 to D1
            sheet.Range( 1, 3, 1, 4 ).SetAutoFilter();

            book.SaveAs( path );
        }

        private void RecursiveLoop2TestMergedInputWorkbook(string path)
        {
            using var book = new XLWorkbook();
            var sheet = book.AddWorksheet( "Sheet1" );

            // Header row (literal — outside the loop).
            sheet.Cell( 1, 2 ).SetValue( "$Name" );

            // 2-row #LoopRow block (insert mode). Row 2 is the directive, row 3 is the
            // block's second row.
            sheet.Column( 1 ).Width = 30;
            sheet.Cell( 2, 1 ).SetValue( "#LoopRow($Loop1, x, 3)" );
            sheet.Range( 2, 2, 2, 4 ).Merge();
            sheet.Cell( 2, 2 ).SetValue( "$x.Text" );
            sheet.Cell( 3, 1 ).SetValue( "#LoopRow($x.Loop2, y, 2)" );
            sheet.Range( 3, 2, 4, 6 ).Merge();
            sheet.Cell( 3, 2 ).SetValue( "$y.Id" );

            book.SaveAs( path );
        }

        private void RecursiveLoop2TestManyRowsAndPropsInputWorkbook(string path) {
            using var book = new XLWorkbook();
            var sheet = book.AddWorksheet( "Sheet1" );

            sheet.Column( 1 ).Width = 30;
            sheet.Cell( 1, 1 ).SetValue( "#LoopRow($Loop1Data, x, 3)" );
            sheet.Range( 1, 2, 1, 43 ).Merge().Style.Fill.BackgroundColor = XLColor.Yellow;
            sheet.Cell( 1, 2 ).SetValue( "$x.Header1" );
            
            sheet.Cell( 2, 1 ).SetValue( "#LoopRow($x.Loop2Data, y, 2)" );
            sheet.Cell( 2, 3 ).SetValue( "$y.Header2" );

            sheet.Cell( 3, 1 ).SetValue( "#LoopRow($y.Loop3Data, z, 1)" );

            System.Reflection.PropertyInfo[] props = typeof( DataManyProps ).GetProperties();
            for (int i = 0; i < props.Length; i++)
            {
                System.Reflection.PropertyInfo pi = props[i];
                sheet.Cell( 3, 4+i ).SetValue( $"$z.{pi.Name}" )
                    .Style.Fill.SetBackgroundColor(
                        (i % 3) switch
                        {
                            0 => XLColor.LightBlue,
                            1 => XLColor.LightGreen,
                            _ => XLColor.LightPink
                        }
                    );
            }

            book.SaveAs( path );
        }

        static void CreateEmptySheetInputWorkbook(string path)
        {
            using var book = new XLWorkbook();
            var sheet = book.AddWorksheet("Sheet1");
            sheet.Cell(1, 1).SetValue("Header");
            sheet.Cell(2, 1).SetValue("#LoopRow($Details, x, 1)");
            sheet.Cell(2, 2).SetValue("$x.Text");

            // Templates created when Excel defaulted to 3 sheets often keep Sheet2/Sheet3
            // with no used cells at all.
            book.AddWorksheet("Sheet2");
            book.AddWorksheet("Sheet3");

            book.SaveAs(path);
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
        public async Task Test2()
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
            using (var stream = new FileStream(Path.Combine(TestEnvironment.PdfSrcPath, "Quotation2.xlsx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var book = new XLWorkbook(stream))
            {
                await book.Worksheet(1).OverWrite(new ObjectExcelSymbolConverter(data));
                book.SaveAs(Path.Combine(TestEnvironment.TestResultsPath, "QuotationDst2.xlsx"));

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
                var lastData = sheet.Cell(31, 22).Value.GetNumber().ToString();
                lastData.Is("106700");


                // R14:Merging cells, Check if it is the same as the value in V26
                var total = sheet.Cell(14, 18).Value.GetNumber().ToString();
                total.Is("106700");

            }

            using var outStream = ExcelConverter.ConvertToPdf(Path.Combine(TestEnvironment.TestResultsPath, "QuotationDst2.xlsx"), 1);
            File.WriteAllBytes(Path.Combine(TestEnvironment.TestResultsPath, "Quotation2.pdf"), outStream.ToArray());
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

                // A1:The Loop directives should not be output as they are, and the cell should be empty.
                sheet.Cell( 1, 1 ).Value.IsBlank.IsTrue();
                sheet.Cell( 2, 1 ).Value.IsBlank.IsTrue();

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

            using (var stream = new FileStream(Path.Combine(TestEnvironment.PdfSrcPath, RecursiveLoop2TestInputFileName), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var book = new XLWorkbook(stream))
            {
                var sheet = book.Worksheets.First();

                await sheet.OverWrite(new ObjectExcelSymbolConverter(data));
                book.SaveAs(Path.Combine(TestEnvironment.TestResultsPath, RecursiveLoop2TestInputFileName ) );

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

                // D2: Fill should be yellow
                var d2FillColor = sheet.Cell( 2, 4 ).Style.Fill.BackgroundColor.Color;
                d2FillColor.Is( XLColor.Yellow.Color, "D2 should be yellow" );

                // D3: Fill should be red
                var d3FillColor = sheet.Cell( 3, 4 ).Style.Fill.BackgroundColor.Color;
                d3FillColor.Is( XLColor.Red.Color, "D3 should be red" );

                // D6: Fill should be yellow
                var d6FillColor = sheet.Cell( 6, 4 ).Style.Fill.BackgroundColor.Color;
                d6FillColor.Is( XLColor.Yellow.Color, "D6 should be yellow" );

                // C1 to D1: should have autofilter
                var autoFilter = sheet.AutoFilter;
                autoFilter.IsNotNull();
                autoFilter.Range.RangeAddress.FirstAddress.RowNumber.Is( 1 );
                autoFilter.Range.RangeAddress.FirstAddress.ColumnNumber.Is( 3 );
                autoFilter.Range.RangeAddress.LastAddress.RowNumber.Is( 1 );
                autoFilter.Range.RangeAddress.LastAddress.ColumnNumber.Is( 4 );

            }

            using var outStream = ExcelConverter.ConvertToPdf(Path.Combine(TestEnvironment.TestResultsPath, RecursiveLoop2TestInputFileName ), 1);
            File.WriteAllBytes(Path.Combine(TestEnvironment.TestResultsPath, Path.ChangeExtension( RecursiveLoop2TestInputFileName, "pdf" ) ), outStream.ToArray());
        }

        [Test]
        public async Task RecursiveLoop2MergedTest()
        {
            var data = new Data {
                Name = "NameA"
            };

            var loop1_1 = new Loop1 {
                Text = "Test1"
            };

            loop1_1.Loop2.Add( new Loop2 { Id = 1 } );
            loop1_1.Loop2.Add( new Loop2 { Id = 2 } );
            loop1_1.Loop2.Add( new Loop2 { Id = 3 } );

            var loop1_2 = new Loop1 {
                Text = "Test2"
            };

            loop1_2.Loop2.Add( new Loop2 { Id = 11 } );
            loop1_2.Loop2.Add( new Loop2 { Id = 22 } );
            loop1_2.Loop2.Add( new Loop2 { Id = 33 } );

            data.Loop1.Add( loop1_1 );
            data.Loop1.Add( loop1_2 );

            using (var stream = new FileStream( Path.Combine( TestEnvironment.PdfSrcPath, RecursiveLoop2TestMergedInputFileName ), FileMode.Open, FileAccess.Read, FileShare.ReadWrite ))
            using (var book = new XLWorkbook( stream )) {
                var sheet = book.Worksheets.First();

                await sheet.OverWrite( new ObjectExcelSymbolConverter( data ) );
                book.SaveAs( Path.Combine( TestEnvironment.TestResultsPath, RecursiveLoop2TestMergedInputFileName ) );

                // B2: should be merged with D2
                var b2MergedRange = sheet.Cell( 2, 2 ).MergedRange();
                b2MergedRange.IsNotNull();
                b2MergedRange.RangeAddress.FirstAddress.RowNumber.Is( 2 );
                b2MergedRange.RangeAddress.FirstAddress.ColumnNumber.Is( 2 );
                b2MergedRange.RangeAddress.LastAddress.RowNumber.Is( 2 );
                b2MergedRange.RangeAddress.LastAddress.ColumnNumber.Is( 4 );

                // B9: should be merged with D9
                var b9MergedRange = sheet.Cell( 9, 2 ).MergedRange();
                b9MergedRange.IsNotNull();
                b9MergedRange.RangeAddress.FirstAddress.RowNumber.Is( 9 );
                b9MergedRange.RangeAddress.FirstAddress.ColumnNumber.Is( 2 );
                b9MergedRange.RangeAddress.LastAddress.RowNumber.Is( 9 );
                b9MergedRange.RangeAddress.LastAddress.ColumnNumber.Is( 4 );

                // B3: should be merged until F4
                var b3MergedRange = sheet.Cell( 3, 2 ).MergedRange();
                b3MergedRange.IsNotNull();
                b3MergedRange.RangeAddress.FirstAddress.RowNumber.Is( 3 );
                b3MergedRange.RangeAddress.FirstAddress.ColumnNumber.Is( 2 );
                b3MergedRange.RangeAddress.LastAddress.RowNumber.Is( 4 );
                b3MergedRange.RangeAddress.LastAddress.ColumnNumber.Is( 6 );

                // B14: should be merged until F15
                var b14MergedRange = sheet.Cell( 14, 2 ).MergedRange();
                b14MergedRange.IsNotNull();
                b14MergedRange.RangeAddress.FirstAddress.RowNumber.Is( 14 );
                b14MergedRange.RangeAddress.FirstAddress.ColumnNumber.Is( 2 );
                b14MergedRange.RangeAddress.LastAddress.RowNumber.Is( 15 );
                b14MergedRange.RangeAddress.LastAddress.ColumnNumber.Is( 6 );
            }

            using var outStream = ExcelConverter.ConvertToPdf( Path.Combine( TestEnvironment.TestResultsPath, RecursiveLoop2TestMergedInputFileName ), 1 );
            File.WriteAllBytes( Path.Combine( TestEnvironment.TestResultsPath, Path.ChangeExtension( RecursiveLoop2TestMergedInputFileName, "pdf" ) ), outStream.ToArray() );
        }

        [Test]
        [Explicit( "This is a test for large row count with many properties, it may take a long time to run" )]
        public async Task RecursiveLoop2LargeRowCountWithManyPropsTest()
        {

            var sourceData = Enumerable
                .Range( 0, 50000 )
                .Select( r =>
                    {
                        var g1Header = string.Concat( "[G1.", r % 100, "] Prop1" );
                        var g2Header = string.Concat( "[G2.", r % 10, "] Prop2" );
                        return new DataManyProps()
                            {
                                Prop1 = g1Header,
                                Prop2 = g2Header,
                                Prop3 = Guid.NewGuid().ToString("N"),
                        }
                        ;
                    }
                )
                .ToArray()
            ;

            var data = new
            {
                Loop1Data = sourceData
                .GroupBy( x => x.Prop1 )
                .Select( g => new
                    {
                        Header1 = g.Key,
                        Loop2Data = g
                            .GroupBy( x => x.Prop2 )
                            .Select( g2 => new
                                {
                                    Header2 = g2.Key,
                                    Loop3Data = g2
                                }
                            )
                            .ToArray()
                    }
                )
                .ToArray()
            };

            var converter = new ObjectExcelSymbolConverter( data );

            byte[] templateBytes;
            using (var templateStream = new FileStream( Path.Combine( TestEnvironment.PdfSrcPath, RecursiveLoop2TestManyRowsAndPropsInputFileName ), FileMode.Open, FileAccess.Read, FileShare.ReadWrite ))
            using (var memoryStream = new MemoryStream())
            {
                templateStream.CopyTo( memoryStream );
                templateBytes = memoryStream.ToArray();
            }

            // Warm up JIT, getter compilation, and ClosedXML before collecting timings.
            using (var warmupBook = new XLWorkbook( new MemoryStream( templateBytes, writable: false ) ))
            {
                await warmupBook.Worksheets.First().OverWrite( converter );
            }

            const int measuredRuns = 3; // Median reduces noise from occasional GC or OS activity.
            var timings = new long[measuredRuns];
            XLWorkbook? lastBook = null;
            var isBookSaved = false;
            try
            {
                for (int i = 0; i < measuredRuns; i++)
                {
                    var book = new XLWorkbook( new MemoryStream( templateBytes, writable: false ) );
                    lastBook?.Dispose();
                    lastBook = book;

                    var sw = System.Diagnostics.Stopwatch.StartNew();
                    await book.Worksheets.First().OverWrite( converter );
                    sw.Stop();
                    timings[i] = sw.ElapsedMilliseconds;

                    if (!isBookSaved) {
                        book.SaveAs( Path.Combine( TestEnvironment.TestResultsPath, RecursiveLoop2TestManyRowsAndPropsInputFileName ) );
                        isBookSaved = true;
                    }
                }

                Array.Sort( timings );
                Console.WriteLine( $"OverWrite timings: {string.Join( ", ", timings )} ms; median: {timings[measuredRuns / 2]} ms" );

                var sheet = lastBook!.Worksheets.First();
                sheet.Cell( 1, 2 ).GetText().Is( "[G1.0] Prop1" );
                sheet.Cell( 2, 3 ).GetText().Is( "[G2.0] Prop2" );
                sheet.Cell( 3, 4 ).GetText().Is( "[G1.0] Prop1" );
                sheet.Cell( 3, 43 ).GetText().Is( "Prop40" );

                // Verify the styled path kept both the repeated header style and alternating leaf styles.
                sheet.Cell( 1, 2 ).Style.Fill.BackgroundColor.Color.Is( XLColor.Yellow.Color, "The repeated header should stay yellow" );
                sheet.Cell( 3, 4 ).Style.Fill.BackgroundColor.Color.Is( XLColor.LightBlue.Color, "The first leaf should stay light blue" );
                sheet.Cell( 3, 5 ).Style.Fill.BackgroundColor.Color.Is( XLColor.LightGreen.Color, "The second leaf should stay light green" );
                sheet.Cell( 1, 2 ).MergedRange().IsNotNull( "The repeated header merge should be restored" );

            }
            finally
            {
                lastBook?.Dispose();
            }

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
            public byte[] Bin { get; set; } = [];
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
                book.Worksheet("last").Cell(2, 2).Value.Is(Blank.Value);
                book.Worksheet("last").Cell(3, 2).Value.Is("★");
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
                book.Worksheet("last").Cell(2, 2).Value.Is(Blank.Value);
                book.Worksheet("last").Cell(3, 2).Value.Is("★");
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
            }

            using var outStream = ExcelConverter.ConvertToPdf(Path.Combine(TestEnvironment.TestResultsPath, "MultiPageSheetBodyLastTest2.xlsx"));
            File.WriteAllBytes(Path.Combine(TestEnvironment.TestResultsPath, "MultiPageSheetBodyLastTest2.pdf"), outStream.ToArray());
        }

        [Test]
        public async Task MultiPageSheetBodyTest()
        {
            var data = new SimpleDataOwner();

            for (int i = 0; i < 90; i++)
            {
                data.Details.Add(new SimpleData { Text = $"Test{i + 1}", Number = i + 1 });
            }

            using (var stream = new FileStream(Path.Combine(TestEnvironment.PdfSrcPath, "MultiPageSheetBodyTest.xlsx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var book = new XLWorkbook(stream))
            {
                await book.OverWrite(new ObjectExcelSymbolConverter(data));
                book.SaveAs(Path.Combine(TestEnvironment.TestResultsPath, "MultiPageSheetBodyTest.xlsx"));

                book.Worksheets.Count.Is(3);
                book.Worksheet("body_0").Cell(2, 2).Value.Is("Test1");
                book.Worksheet("body_0").Cell(31, 2).Value.Is("Test30");
            }

            using var outStream = ExcelConverter.ConvertToPdf(Path.Combine(TestEnvironment.TestResultsPath, "MultiPageSheetBodyTest.xlsx"));
            File.WriteAllBytes(Path.Combine(TestEnvironment.TestResultsPath, "MultiPageSheetBodyTest.pdf"), outStream.ToArray());
        }

        [Test]
        public async Task MultiPageSheetPageTest()
        {
            var data = new SimpleDataOwner();

            for (int i = 0; i < 100; i++)
            {
                data.Details.Add(new SimpleData { Text = $"Test{i + 1}", Number = i + 1 });
            }

            using (var stream = new FileStream(Path.Combine(TestEnvironment.PdfSrcPath, "MultiPageSheetPageTest.xlsx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var book = new XLWorkbook(stream))
            {
                await book.OverWrite(new ObjectExcelSymbolConverter(data));
                book.SaveAs(Path.Combine(TestEnvironment.TestResultsPath, "MultiPageSheetPageTest.xlsx"));
            }

            using var outStream = ExcelConverter.ConvertToPdf(Path.Combine(TestEnvironment.TestResultsPath, "MultiPageSheetPageTest.xlsx"));
            File.WriteAllBytes(Path.Combine(TestEnvironment.TestResultsPath, "MultiPageSheetPageTest.pdf"), outStream.ToArray());
        }

        [Test]
        public async Task EmptySheetTest()
        {
            var data = new SimpleDataOwner();
            data.Details.Add(new SimpleData { Text = "Test1", Number = 1 });
            data.Details.Add(new SimpleData { Text = "Test2", Number = 2 });

            using (var stream = new FileStream(Path.Combine(TestEnvironment.PdfSrcPath, EmptySheetInputFileName), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var book = new XLWorkbook(stream))
            {
                // Must not throw even though Sheet2/Sheet3 have no used cells.
                await book.OverWrite(new ObjectExcelSymbolConverter(data));
                book.SaveAs(Path.Combine(TestEnvironment.TestResultsPath, "EmptySheetTest.xlsx"));

                var sheet = book.Worksheet("Sheet1");
                sheet.Cell(1, 1).Value.GetText().Is("Header");
                sheet.Cell(2, 2).Value.GetText().Is("Test1");
                sheet.Cell(3, 2).Value.GetText().Is("Test2");
            }
        }

        [Test]
        public async Task QRCode()
        {
            var data = new SimpleData { Text = "https://www.codeer.co.jp/" };

            using (var stream = new FileStream(Path.Combine(TestEnvironment.PdfSrcPath, "QR.xlsx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var book = new XLWorkbook(stream))
            {
                await book.OverWrite(new ObjectExcelSymbolConverter(data));
                book.SaveAs(Path.Combine(TestEnvironment.TestResultsPath, "QR.xlsx"));
            }

            using var outStream = ExcelConverter.ConvertToPdf(Path.Combine(TestEnvironment.TestResultsPath, "QR.xlsx"));
            File.WriteAllBytes(Path.Combine(TestEnvironment.TestResultsPath, "QR.pdf"), outStream.ToArray());
        }

        [Test]
        public async Task Image()
        {
            var data = new SimpleData { Bin = Resources.ImageSample };

            using (var stream = new FileStream(Path.Combine(TestEnvironment.PdfSrcPath, "Image.xlsx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var book = new XLWorkbook(stream))
            {
                await book.OverWrite(new ObjectExcelSymbolConverter(data));
                book.SaveAs(Path.Combine(TestEnvironment.TestResultsPath, "Image.xlsx"));
            }

            using var outStream = ExcelConverter.ConvertToPdf(Path.Combine(TestEnvironment.TestResultsPath, "Image.xlsx"));
            File.WriteAllBytes(Path.Combine(TestEnvironment.TestResultsPath, "Image.pdf"), outStream.ToArray());
        }

        [Test]
        public async Task ImageQRLoop()
        {
            var data = new SimpleDataOwner();

            data.Details.Add(new SimpleData { Text = "https://www.codeer.co.jp/", Bin = Resources.ImageSample });

            using (var stream = new FileStream(Path.Combine(TestEnvironment.PdfSrcPath, "ImageQRLoop.xlsx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var book = new XLWorkbook(stream))
            {
                await book.OverWrite(new ObjectExcelSymbolConverter(data));
                book.SaveAs(Path.Combine(TestEnvironment.TestResultsPath, "ImageQRLoop.xlsx"));
            }

            using var outStream = ExcelConverter.ConvertToPdf(Path.Combine(TestEnvironment.TestResultsPath, "ImageQRLoop.xlsx"));
            File.WriteAllBytes(Path.Combine(TestEnvironment.TestResultsPath, "ImageQRLoop.pdf"), outStream.ToArray());
        }
    }
}
