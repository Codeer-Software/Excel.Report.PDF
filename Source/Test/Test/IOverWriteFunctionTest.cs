using ClosedXML.Excel;
using Excel.Report.PDF;
using NUnit.Framework;

namespace Test
{
    public class IOverWriteFunctionTest
    {
        class TestItem
        {
            public string Label { get; set; } = string.Empty;
            public int Number { get; set; }
        }

        class TestData
        {
            public string Name { get; set; } = string.Empty;
            public List<TestItem> Items { get; set; } = new();
        }

        // The IOverWriteFunction under test. Joins all args with '-' and writes them to the cell.
        // Mirrors the built-in Image/QR functions: bails when its primary arg can't be resolved,
        // which happens on the pre-loop pass for cells inside a #LoopRow block.
        class JoinOverWriteFunction : IOverWriteFunction
        {
            public string Name => "TestJoin";

            public List<(int Row, int Col, object?[] Args)> Invocations { get; } = new();

            public Task InvokeAsync(IXLWorksheet sheet, int rowIndex, int colIndex, object?[] args)
            {
                if (args.Length == 0 || args[0] == null) return Task.CompletedTask;

                Invocations.Add((rowIndex, colIndex, args));
                var text = string.Join("-", args.Select(a => a?.ToString() ?? string.Empty));
                sheet.Cell(rowIndex, colIndex).SetValue(XLCellValue.FromObject(text));
                return Task.CompletedTask;
            }
        }

        // A "naive" IOverWriteFunction: no bail-out for unresolved args. This is what most users
        // would write on a first attempt — and it surfaces the pre-loop-pass clobbering bug when
        // placed inside a #LoopRow block.
        class NaiveOverWriteFunction : IOverWriteFunction
        {
            public string Name => "TestNaive";

            public Task InvokeAsync(IXLWorksheet sheet, int rowIndex, int colIndex, object?[] args)
            {
                var text = string.Join("/", args.Select(a => a?.ToString() ?? string.Empty));
                sheet.Cell(rowIndex, colIndex).SetValue(XLCellValue.FromObject(text));
                return Task.CompletedTask;
            }
        }

        static readonly JoinOverWriteFunction _joinFunction = new();
        static readonly NaiveOverWriteFunction _naiveFunction = new();
        static bool _registered;

        const string InputFileName = "IOverWriteFunctionTest.xlsx";
        const string LoopInputFileName = "IOverWriteFunctionTest_Loop.xlsx";
        const string LoopRootRefInputFileName = "IOverWriteFunctionTest_LoopRootRef.xlsx";
        const string LoopRowDataInputFileName = "IOverWriteFunctionTest_LoopRowData.xlsx";
        const string MultiRowBlockInputFileName = "IOverWriteFunctionTest_MultiRowBlock.xlsx";

        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            if (!Directory.Exists(TestEnvironment.TestResultsPath))
            {
                Directory.CreateDirectory(TestEnvironment.TestResultsPath);
            }

            // Custom function registration is global; register only once even if tests rerun.
            if (!_registered)
            {
                ExcelOverWriter.RegisterOverWriteFunction(_joinFunction);
                ExcelOverWriter.RegisterOverWriteFunction(_naiveFunction);
                _registered = true;
            }

            // Generate the input xlsx into the Data folder if it does not exist yet.
            var inputPath = Path.Combine(TestEnvironment.PdfSrcPath, InputFileName);
            if (!File.Exists(inputPath))
            {
                Directory.CreateDirectory(TestEnvironment.PdfSrcPath);
                CreateInputWorkbook(inputPath);
            }

            var loopInputPath = Path.Combine(TestEnvironment.PdfSrcPath, LoopInputFileName);
            if (!File.Exists(loopInputPath))
            {
                Directory.CreateDirectory(TestEnvironment.PdfSrcPath);
                CreateLoopInputWorkbook(loopInputPath);
            }

            var loopRootRefPath = Path.Combine(TestEnvironment.PdfSrcPath, LoopRootRefInputFileName);
            if (!File.Exists(loopRootRefPath))
            {
                Directory.CreateDirectory(TestEnvironment.PdfSrcPath);
                CreateLoopRootRefInputWorkbook(loopRootRefPath);
            }

            var loopRowDataPath = Path.Combine(TestEnvironment.PdfSrcPath, LoopRowDataInputFileName);
            if (!File.Exists(loopRowDataPath))
            {
                Directory.CreateDirectory(TestEnvironment.PdfSrcPath);
                CreateLoopRowDataInputWorkbook(loopRowDataPath);
            }

            var multiRowBlockPath = Path.Combine(TestEnvironment.PdfSrcPath, MultiRowBlockInputFileName);
            if (!File.Exists(multiRowBlockPath))
            {
                Directory.CreateDirectory(TestEnvironment.PdfSrcPath);
                CreateMultiRowBlockInputWorkbook(multiRowBlockPath);
            }
        }

        static void CreateLoopRowDataInputWorkbook(string path)
        {
            using var book = new XLWorkbook();
            var sheet = book.AddWorksheet("Sheet1");

            // Header.
            sheet.Cell(1, 1).SetValue("Label");
            sheet.Cell(1, 2).SetValue("Func");

            // #LoopRowData directive (no row-insert; pre-allocated rows below).
            sheet.Cell(2, 1).SetValue("#LoopRowData($Items, item)");
            sheet.Cell(2, 2).SetValue("#TestNaive($item.Label, $item.Number)");

            // LoopRowData does NOT insert rows — it writes into rows that already exist below
            // the template. ClosedXML auto-creates rows on cell access, so we don't need to
            // pre-populate them here.

            book.SaveAs(path);
        }

        static void CreateMultiRowBlockInputWorkbook(string path)
        {
            using var book = new XLWorkbook();
            var sheet = book.AddWorksheet("Sheet1");

            // Header row.
            sheet.Cell(1, 1).SetValue("Header");

            // 2-row block: row 2 is the directive (with $item.Label), row 3 has the custom
            // function and a numeric $-reference. Functions on the non-directive row of the
            // block should NOT be suppressed — they must be invoked in the recursive pass.
            sheet.Cell(2, 1).SetValue("#LoopRow($Items, item, 2)");
            sheet.Cell(2, 2).SetValue("$item.Label");
            sheet.Cell(3, 3).SetValue("#TestNaive($item.Label, $item.Number)");
            sheet.Cell(3, 4).SetValue("$item.Number");

            book.SaveAs(path);
        }

        static void CreateLoopRootRefInputWorkbook(string path)
        {
            using var book = new XLWorkbook();
            var sheet = book.AddWorksheet("Sheet1");

            // Header row (literal — outside the loop).
            sheet.Cell(1, 1).SetValue("Owner");
            sheet.Cell(1, 2).SetValue("Label");
            sheet.Cell(1, 3).SetValue("Tag");

            // Loop directive row mixes:
            //   - A: the loop directive itself
            //   - B: a root-level $-reference ($Name) — needs to be resolved by the outer pass
            //   - C: a per-iteration $-reference ($item.Label) — needs the element converter
            //   - D: a custom function reading per-iteration values
            sheet.Cell(2, 1).SetValue("#LoopRow($Items, item)");
            sheet.Cell(2, 2).SetValue("$Name");
            sheet.Cell(2, 3).SetValue("$item.Label");
            sheet.Cell(2, 4).SetValue("#TestNaive($item.Label, $item.Number)");

            book.SaveAs(path);
        }

        static void CreateLoopInputWorkbook(string path)
        {
            using var book = new XLWorkbook();
            var sheet = book.AddWorksheet("Sheet1");

            // Header row (literal — outside the loop).
            sheet.Cell(1, 1).SetValue("Label");
            sheet.Cell(1, 2).SetValue("Number");

            // Loop block. The custom function template lives in column B of the looped row.
            sheet.Cell(2, 1).SetValue("#LoopRow($Items, item)");
            sheet.Cell(2, 2).SetValue("#TestNaive($item.Label, $item.Number)");

            book.SaveAs(path);
        }

        static void CreateInputWorkbook(string path)
        {
            using var book = new XLWorkbook();
            var sheet = book.AddWorksheet("Sheet1");

            // Plain $-substitution (sanity baseline, not the function under test).
            sheet.Cell(1, 1).SetValue("Name:");
            sheet.Cell(1, 2).SetValue("$Name");

            // Custom function with two literal args.
            sheet.Cell(2, 1).SetValue("Literal:");
            sheet.Cell(2, 2).SetValue("#TestJoin(Hello, World)");

            // Custom function mixing a $-reference and a literal arg.
            sheet.Cell(3, 1).SetValue("Mixed:");
            sheet.Cell(3, 2).SetValue("#TestJoin($Name, fixed)");

            // Custom function with a single arg.
            sheet.Cell(4, 1).SetValue("Single:");
            sheet.Cell(4, 2).SetValue("#TestJoin(only)");

            // Inside a loop: custom function resolves $item.* references for each iteration.
            sheet.Cell(5, 1).SetValue("#LoopRow($Items, item)");
            sheet.Cell(5, 2).SetValue("#TestJoin($item.Label, $item.Number)");

            book.SaveAs(path);
        }

        [Test]
        public async Task InvokesCustomFunctionWithLiteralAndReferenceArgs()
        {
            _joinFunction.Invocations.Clear();

            var data = new TestData
            {
                Name = "Tatsuya",
                Items =
                {
                    new TestItem { Label = "A", Number = 1 },
                    new TestItem { Label = "B", Number = 2 },
                    new TestItem { Label = "C", Number = 3 },
                }
            };

            var inputPath = Path.Combine(TestEnvironment.PdfSrcPath, InputFileName);
            var outputPath = Path.Combine(TestEnvironment.TestResultsPath, "IOverWriteFunctionTest.xlsx");

            using (var stream = new FileStream(inputPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var book = new XLWorkbook(stream))
            {
                await book.Worksheet(1).OverWrite(new ObjectExcelSymbolConverter(data));
                book.SaveAs(outputPath);

                var sheet = book.Worksheets.First();

                // Baseline: $Name was replaced as plain text substitution.
                sheet.Cell(1, 2).Value.GetText().Is("Tatsuya");

                // Two literal args: both should arrive verbatim.
                sheet.Cell(2, 2).Value.GetText().Is("Hello-World");

                // $-reference resolved + literal preserved.
                sheet.Cell(3, 2).Value.GetText().Is("Tatsuya-fixed");

                // Single arg call still works.
                sheet.Cell(4, 2).Value.GetText().Is("only");

                // Loop expansion: row 5/6/7 each invoked the function with the item's properties.
                sheet.Cell(5, 2).Value.GetText().Is("A-1");
                sheet.Cell(6, 2).Value.GetText().Is("B-2");
                sheet.Cell(7, 2).Value.GetText().Is("C-3");
            }

            // Invocations: rows 2/3/4 from the non-loop section + rows 5/6/7 from the loop expansion = 6.
            // The pre-loop pass on row 5 skips function invocation entirely (loop-row suppression),
            // so no extra "args were null" call is recorded.
            _joinFunction.Invocations.Count.Is(6);

            // Row/col arrive correctly. Verify the literal call at (2, 2).
            var literalCall = _joinFunction.Invocations[0];
            literalCall.Row.Is(2);
            literalCall.Col.Is(2);
            literalCall.Args.Length.Is(2);
            literalCall.Args[0]!.ToString().Is("Hello");
            literalCall.Args[1]!.ToString().Is("World");

            // $-reference args are resolved to their underlying values.
            var mixedCall = _joinFunction.Invocations[1];
            mixedCall.Args[0].Is("Tatsuya");
            mixedCall.Args[1]!.ToString().Is("fixed");

            // In the loop, $item.Number arrives as the boxed int (not a string).
            var firstLoopCall = _joinFunction.Invocations[3];
            firstLoopCall.Row.Is(5);
            firstLoopCall.Args[0].Is("A");
            firstLoopCall.Args[1].Is(1);
        }

        // Expected behavior: a custom IOverWriteFunction placed inside a #LoopRow block should
        // be re-evaluated once per loop iteration, with the iteration's $item.* references
        // resolved against the per-element converter.
        //
        // Actual (buggy) behavior: OverWriteCell runs on the #LoopRow source row BEFORE the loop
        // is detected, so the function is invoked once with the root converter — where $item.*
        // resolves to null — and the function clobbers the template cell. CopyRows then duplicates
        // the already-clobbered cell, and the recursive pass finds nothing to rewrite, so every
        // expanded row ends up holding whatever the first (unresolved) invocation produced.
        [Test]
        public async Task CustomFunctionInsideLoopRow_ExpandsPerIteration()
        {
            var data = new TestData
            {
                Items =
                {
                    new TestItem { Label = "A", Number = 1 },
                    new TestItem { Label = "B", Number = 2 },
                    new TestItem { Label = "C", Number = 3 },
                }
            };

            var inputPath = Path.Combine(TestEnvironment.PdfSrcPath, LoopInputFileName);
            var outputPath = Path.Combine(TestEnvironment.TestResultsPath, "IOverWriteFunctionTest_Loop.xlsx");

            using var stream = new FileStream(inputPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var book = new XLWorkbook(stream);

            await book.Worksheet(1).OverWrite(new ObjectExcelSymbolConverter(data));
            book.SaveAs(outputPath);

            var sheet = book.Worksheets.First();

            // Header row is untouched.
            sheet.Cell(1, 1).Value.GetText().Is("Label");
            sheet.Cell(1, 2).Value.GetText().Is("Number");

            // Each iteration's $item.* should have been resolved into the function call.
            sheet.Cell(2, 2).Value.GetText().Is("A/1");
            sheet.Cell(3, 2).Value.GetText().Is("B/2");
            sheet.Cell(4, 2).Value.GetText().Is("C/3");
        }

        // A #LoopRow row may also contain root-level $-references (e.g. $Name) and per-iteration
        // $-references (e.g. $item.Label) on the same row. The root-level reference must be
        // resolved by the outer pass (the per-element converter doesn't know about root props),
        // while the per-iteration reference is resolved by the recursive pass. Both must coexist
        // with a custom function on the same row.
        [Test]
        public async Task LoopDirectiveRow_ResolvesBothRootAndItemReferences()
        {
            var data = new TestData
            {
                Name = "Tatsuya",
                Items =
                {
                    new TestItem { Label = "A", Number = 1 },
                    new TestItem { Label = "B", Number = 2 },
                    new TestItem { Label = "C", Number = 3 },
                }
            };

            var inputPath = Path.Combine(TestEnvironment.PdfSrcPath, LoopRootRefInputFileName);
            var outputPath = Path.Combine(TestEnvironment.TestResultsPath, "IOverWriteFunctionTest_LoopRootRef.xlsx");

            using var stream = new FileStream(inputPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var book = new XLWorkbook(stream);

            await book.Worksheet(1).OverWrite(new ObjectExcelSymbolConverter(data));
            book.SaveAs(outputPath);

            var sheet = book.Worksheets.First();

            // Header row untouched.
            sheet.Cell(1, 1).Value.GetText().Is("Owner");
            sheet.Cell(1, 2).Value.GetText().Is("Label");
            sheet.Cell(1, 3).Value.GetText().Is("Tag");

            // Root-level $Name must be filled on every expanded row.
            sheet.Cell(2, 2).Value.GetText().Is("Tatsuya");
            sheet.Cell(3, 2).Value.GetText().Is("Tatsuya");
            sheet.Cell(4, 2).Value.GetText().Is("Tatsuya");

            // Per-iteration $item.Label resolves to each element.
            sheet.Cell(2, 3).Value.GetText().Is("A");
            sheet.Cell(3, 3).Value.GetText().Is("B");
            sheet.Cell(4, 3).Value.GetText().Is("C");

            // Custom function runs once per iteration with the element converter.
            sheet.Cell(2, 4).Value.GetText().Is("A/1");
            sheet.Cell(3, 4).Value.GetText().Is("B/2");
            sheet.Cell(4, 4).Value.GetText().Is("C/3");
        }

        // #LoopRowData has IsInsertMode=false and IsFormatCopy=false — a different CopyRows
        // path than #LoopRow. The fix's loop-row detection uses StartsWith("#LoopRow") which
        // must also match "#LoopRowData", so the function-suppression on the directive row
        // still applies.
        [Test]
        public async Task CustomFunctionInsideLoopRowData_ExpandsPerIteration()
        {
            var data = new TestData
            {
                Items =
                {
                    new TestItem { Label = "A", Number = 1 },
                    new TestItem { Label = "B", Number = 2 },
                    new TestItem { Label = "C", Number = 3 },
                }
            };

            var inputPath = Path.Combine(TestEnvironment.PdfSrcPath, LoopRowDataInputFileName);
            var outputPath = Path.Combine(TestEnvironment.TestResultsPath, "IOverWriteFunctionTest_LoopRowData.xlsx");

            using var stream = new FileStream(inputPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var book = new XLWorkbook(stream);

            await book.Worksheet(1).OverWrite(new ObjectExcelSymbolConverter(data));
            book.SaveAs(outputPath);

            var sheet = book.Worksheets.First();

            // Each iteration's $item.* should arrive resolved at the function.
            sheet.Cell(2, 2).Value.GetText().Is("A/1");
            sheet.Cell(3, 2).Value.GetText().Is("B/2");
            sheet.Cell(4, 2).Value.GetText().Is("C/3");
        }

        // For a multi-row loop block (rowCopyCount > 1), only the directive row gets function
        // suppression. Functions on the *other* rows of the block must still be invoked
        // normally during the recursive pass. This pins down the boundary of the fix.
        [Test]
        public async Task CustomFunctionOnNonDirectiveRow_OfMultiRowBlock_IsInvoked()
        {
            var data = new TestData
            {
                Items =
                {
                    new TestItem { Label = "A", Number = 1 },
                    new TestItem { Label = "B", Number = 2 },
                    new TestItem { Label = "C", Number = 3 },
                }
            };

            var inputPath = Path.Combine(TestEnvironment.PdfSrcPath, MultiRowBlockInputFileName);
            var outputPath = Path.Combine(TestEnvironment.TestResultsPath, "IOverWriteFunctionTest_MultiRowBlock.xlsx");

            using var stream = new FileStream(inputPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var book = new XLWorkbook(stream);

            await book.Worksheet(1).OverWrite(new ObjectExcelSymbolConverter(data));
            book.SaveAs(outputPath);

            var sheet = book.Worksheets.First();

            // Iter 0: rows 2-3.
            sheet.Cell(2, 2).Value.GetText().Is("A");
            sheet.Cell(3, 3).Value.GetText().Is("A/1");
            sheet.Cell(3, 4).Value.GetNumber().Is(1d);

            // Iter 1: rows 4-5.
            sheet.Cell(4, 2).Value.GetText().Is("B");
            sheet.Cell(5, 3).Value.GetText().Is("B/2");
            sheet.Cell(5, 4).Value.GetNumber().Is(2d);

            // Iter 2: rows 6-7.
            sheet.Cell(6, 2).Value.GetText().Is("C");
            sheet.Cell(7, 3).Value.GetText().Is("C/3");
            sheet.Cell(7, 4).Value.GetNumber().Is(3d);
        }
    }
}
