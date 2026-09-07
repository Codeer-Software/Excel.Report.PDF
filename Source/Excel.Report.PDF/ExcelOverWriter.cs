using System.Collections;
using ClosedXML.Excel;

namespace Excel.Report.PDF
{
    /// <summary>
    /// Expands template loops and replaces symbols in Excel worksheets.
    /// </summary>
    public static class ExcelOverWriter
    {
        #region Public API

        static readonly List<IOverWriteFunction> Functions = new() { new ImageOverWriteFunction(), new QRCodeOverWriteFunction() };

        /// <summary>
        /// Adds a custom function that templates can use.
        /// </summary>
        public static void RegisterOverWriteFunction(IOverWriteFunction function) => RegisterFunction(function);

        /// <summary>
        /// Expands every sheet in a workbook after preparing any paged sheets.
        /// </summary>
        public static async Task OverWrite(this XLWorkbook book, IExcelSymbolConverter converter)
        {
            var pagePlans = await ExcelPageLoopProcessor.BuildPagePlansAsync(book, converter);
            ExcelPageLoopProcessor.MaterializeBodyPageSheets(book, pagePlans);

            // Use a copy because writing a sheet creates a temporary snapshot sheet.
            foreach (var sheet in book.Worksheets.ToList())
                await WriteWorksheetAsync(sheet, converter, pagePlans.Values.ToList());
        }

        /// <summary>
        /// Expands one worksheet without preparing paged sheets.
        /// </summary>
        public static async Task OverWrite(this IXLWorksheet sheet, IExcelSymbolConverter converter)
            => await WriteWorksheetAsync(sheet, converter, Array.Empty<ExcelPageLoopProcessor.PageLoopPlan>());

        /// <summary>
        /// Stores custom functions in one shared list so every worksheet uses the same functions.
        /// </summary>
        static void RegisterFunction(IOverWriteFunction function) => Functions.Add(function);

        #endregion

        #region Worksheet pipeline

        /// <summary>
        /// Coordinates the snapshot, planning, and materialization phases for one worksheet.
        /// Keeping these phases in order prevents row changes from invalidating source coordinates.
        /// </summary>
        static async Task WriteWorksheetAsync(IXLWorksheet sheet, IExcelSymbolConverter converter, IReadOnlyList<ExcelPageLoopProcessor.PageLoopPlan> pagePlans)
        {
            ExcelUtils.GetRowColCount(sheet, out var rowCount, out var colCount);
            var template = SnapshotTemplate(sheet, rowCount);
            var autoFilter = CaptureAutoFilter(sheet);
            var plan = await BuildExpansionPlanAsync(sheet, rowCount, converter, pagePlans, template);

            await MaterializeWorksheetAsync(sheet, template, plan, rowCount, colCount, autoFilter);
        }

        /// <summary>
        /// Writes the planned rows using batch writes or template range copies.
        /// </summary>
        static async Task MaterializeWorksheetAsync(IXLWorksheet sheet, TemplateSnapshot template, ExpansionPlan plan, int sourceRowCount, int colCount, AutoFilterSnapshot? autoFilter)
        {
            if (CanUsePlainValueBatch(sheet, template, plan))
            {
                ReserveWorksheetRows(sheet, plan.Rows.Count, sourceRowCount, colCount);
                await WriteSymbolValuesAsBatchAsync(sheet, plan.Rows, colCount);
                RestoreAutoFilter(sheet, autoFilter);
                return;
            }

            var workbook = sheet.Workbook;
            var sourceName = $"__EOWT_{Guid.NewGuid():N}"[..15]; // Excel limits worksheet names to 31 characters.
            var sourceSheet = sheet.CopyTo(sourceName);
            try
            {
                ReserveWorksheetRows(sheet, plan.Rows.Count, sourceRowCount, colCount);
                if (template.CanUseStyledValueBatch)
                {
                    // Copy styles in grouped ranges and values in one matrix so formatting does not force one CopyTo per row.
                    ApplyStyleGroups(sheet, sourceSheet, plan.Rows, sourceRowCount, colCount);
                    await WriteSymbolValuesAsBatchAsync(sheet, plan.Rows, colCount);
                    RestoreMergedRanges(sheet, plan.Rows, plan.Merges);
                }
                else
                {
                    // Copy template cells to preserve formulas, literal values, and function calls.
                    CopyTemplateRanges(sheet, sourceSheet, plan.Rows, sourceRowCount, colCount);
                    RestoreMergedRanges(sheet, plan.Rows, plan.Merges);
                    await ApplyCellOperationsAsync(sheet, plan.Rows);
                }
                RestoreAutoFilter(sheet, autoFilter);
            }
            finally
            {
                workbook.Worksheets.Delete(sourceName);
            }
        }

        static AutoFilterSnapshot? CaptureAutoFilter(IXLWorksheet sheet)
        {
            var autoFilter = sheet.AutoFilter;
            if (autoFilter == null || !autoFilter.IsEnabled || autoFilter.Range == null) return null;

            var range = autoFilter.Range.RangeAddress;
            return new AutoFilterSnapshot(
                range.FirstAddress.RowNumber,
                range.FirstAddress.ColumnNumber,
                range.LastAddress.RowNumber,
                range.LastAddress.ColumnNumber);
        }

        static void RestoreAutoFilter(IXLWorksheet sheet, AutoFilterSnapshot? snapshot)
        {
            if (snapshot == null) return;
            var value = snapshot.Value;
            sheet.Range(
                value.FirstRow,
                value.FirstColumn,
                value.LastRow,
                value.LastColumn).SetAutoFilter();
        }

        #endregion

        #region Planning

        /// <summary>
        /// Reads template operations once so generated rows can reuse them.
        /// </summary>
        static TemplateSnapshot SnapshotTemplate(IXLWorksheet sheet, int rowCount)
        {
            var rows = new IReadOnlyList<TemplateCellOperation>[rowCount + 1];
            var leftText = new string[rowCount + 1];
            var canUsePlainValueBatch = true;
            var canUseStyledValueBatch = true;
            var defaultStyle = XLWorkbook.DefaultStyle;

            for (var rowNumber = 1; rowNumber <= rowCount; rowNumber++)
            {
                leftText[rowNumber] = sheet.GetText(rowNumber, 1).Trim();
                var operations = new List<TemplateCellOperation>();
                foreach (var cell in sheet.Row(rowNumber).CellsUsed(XLCellsUsedOptions.All))
                {
                    var text = cell.GetString().Trim();
                    var symbol = text.StartsWith("$") ? text.Substring(1) : null;
                    var operationKind = GetOperationKind(text, symbol);
                    var isValueOperation = operationKind == TemplateCellKind.Symbol || operationKind == TemplateCellKind.Directive;

                    // Inspect every used cell, including literals and styled blanks, before filtering operations.
                    if (cell.HasFormula || !cell.Style.Equals(defaultStyle) || !isValueOperation)
                        canUsePlainValueBatch = false;
                    if (cell.HasFormula || (text.Length > 0 && !isValueOperation))
                        canUseStyledValueBatch = false;

                    if (operationKind != TemplateCellKind.Literal)
                        operations.Add(new TemplateCellOperation(cell.Address.ColumnNumber, text, symbol, operationKind));
                }
                rows[rowNumber] = operations.ToArray();
            }

            return new TemplateSnapshot(rows, leftText, canUsePlainValueBatch, canUseStyledValueBatch);
        }

        /// <summary>
        /// Resolves all loops before changing the worksheet.
        /// This keeps source row numbers stable while the plan is built.
        /// </summary>
        static async Task<ExpansionPlan> BuildExpansionPlanAsync(IXLWorksheet sheet, int rowCount, IExcelSymbolConverter converter, IReadOnlyList<ExcelPageLoopProcessor.PageLoopPlan> pagePlans, TemplateSnapshot template)
        {
            var rows = new List<OutputRowPlan>(rowCount);
            var planner = new ExpansionPlanner(sheet.Name, pagePlans, template, rows);
            await planner.PlanRangeAsync(1, rowCount, converter, new[] { converter }, false, RowFormattingMode.CopyTemplate);

            var merges = sheet.MergedRanges.Select(range => new MergeOperation(
                range.RangeAddress.FirstAddress.RowNumber,
                range.RangeAddress.LastAddress.RowNumber,
                range.RangeAddress.FirstAddress.ColumnNumber,
                range.RangeAddress.LastAddress.ColumnNumber)).ToArray();
            return new ExpansionPlan(rows, merges);
        }

        /// <summary>
        /// Holds worksheet-wide planning state shared by every recursive range.
        /// </summary>
        sealed class ExpansionPlanner
        {
            readonly string sheetName;
            readonly IReadOnlyList<ExcelPageLoopProcessor.PageLoopPlan> pagePlans;
            readonly TemplateSnapshot template;
            readonly List<OutputRowPlan> rows;

            public ExpansionPlanner(string sheetName, IReadOnlyList<ExcelPageLoopProcessor.PageLoopPlan> pagePlans, TemplateSnapshot template, List<OutputRowPlan> rows)
            {
                this.sheetName = sheetName;
                this.pagePlans = pagePlans;
                this.template = template;
                this.rows = rows;
            }

            /// <summary>
            /// Adds output rows for nested loops.
            /// Data-only loops also scan their preallocated rows to preserve their existing behavior.
            /// </summary>
            public async Task PlanRangeAsync(int startRow, int endRow, IExcelSymbolConverter converter, IReadOnlyList<IExcelSymbolConverter> converterScopes, bool clearFirstDirective, RowFormattingMode formattingMode)
            {
                var sourceRow = startRow;
                var effectiveEndRow = endRow;
                while (sourceRow <= effectiveEndRow)
                {
                    var leftText = template.GetLeftText(sourceRow);
                    if (clearFirstDirective && sourceRow == startRow)
                    {
                        // The parent owns the marker, but the other cells still run for this item.
                        AddOutputRow(sourceRow, converterScopes, RowCleanup.Directive, formattingMode);
                        sourceRow++;
                        continue;
                    }

                    if (!leftText.StartsWith("#LoopRow") && !leftText.StartsWith("#PagedLoopRows"))
                    {
                        // Only valid loop directives change the row layout; other text is copied normally.
                        AddOutputRow(sourceRow, converterScopes, RowCleanup.None, formattingMode);
                        sourceRow++;
                        continue;
                    }

                    var loop = new LoopPlan();
                    if (!await ParseLoopAsync(leftText, converter, loop, sheetName, pagePlans))
                    {
                        // Invalid directives stay as text for compatibility.
                        AddOutputRow(sourceRow, converterScopes, RowCleanup.None, formattingMode);
                        sourceRow++;
                        continue;
                    }

                    var blockEnd = sourceRow + loop.RowCopyCount - 1;
                    if (loop.Items.Count == 0)
                    {
                        if (loop.Mode == LoopMode.InsertRows)
                        {
                            // Omit the whole block without deleting rows one at a time.
                            sourceRow = blockEnd + 1;
                            continue;
                        }

                        // Data-only loops keep their row, but clear its old symbols and marker.
                        AddOutputRow(sourceRow, converterScopes, RowCleanup.Directive | RowCleanup.Symbols, GetChildFormattingMode(loop.Mode, formattingMode));
                        sourceRow++;
                        continue;
                    }

                    var emittedStart = rows.Count;
                    foreach (var item in loop.Items)
                    {
                        var child = converter.CreateChildExcelSymbolConverter(item, loop.Name);
                        var childScopes = converterScopes.Concat(new[] { child }).ToArray();
                        await PlanRangeAsync(sourceRow, blockEnd, child, childScopes, true, GetChildFormattingMode(loop.Mode, formattingMode));
                    }

                    if (loop.Mode == LoopMode.InsertRows)
                    {
                        // Only the original block is consumed; generated copies exist in the plan.
                        sourceRow = blockEnd + 1;
                    }
                    else
                    {
                        var emittedRows = rows.Count - emittedStart;
                        // Data-only loops consume the rows they fill, including new rows beyond the template.
                        sourceRow += emittedRows;
                        effectiveEndRow += emittedRows - loop.RowCopyCount;
                    }
                }
            }

            /// <summary>
            /// Adds a planned row with its source operations and converter scopes.
            /// The output row number is assigned later during materialization.
            /// </summary>
            void AddOutputRow(int sourceRow, IReadOnlyList<IExcelSymbolConverter> converterScopes, RowCleanup cleanup, RowFormattingMode formattingMode)
                => rows.Add(new OutputRowPlan(sourceRow, converterScopes, cleanup, formattingMode, template.GetRow(sourceRow)));
        }

        /// <summary>
        /// Parses a loop and resolves its collection so the planner knows how many rows to create.
        /// </summary>
        static async Task<bool> ParseLoopAsync(string text, IExcelSymbolConverter converter, LoopPlan loop, string sheetName, IReadOnlyList<ExcelPageLoopProcessor.PageLoopPlan> pagePlans)
        {
            if (text.StartsWith("#LoopRow"))
                return await ParseOrdinaryLoopAsync(text, converter, loop);

            return ParsePagedLoop(text, loop, sheetName, pagePlans);
        }

        static async Task<bool> ParseOrdinaryLoopAsync(string text, IExcelSymbolConverter converter, LoopPlan loop)
        {
            var dataOnly = text.StartsWith("#LoopRowData");
            var prefix = dataOnly ? "#LoopRowData" : "#LoopRow";
            var args = text.Replace(prefix, "").Replace("(", "").Replace(")", "").Split(',').Select(value => value.Trim()).ToArray();
            var rowCopyCount = 1;
            if (args.Length == 3 && !int.TryParse(args[2], out rowCopyCount)) return false;
            if (args.Length < 2 || !args[0].StartsWith("$")) return false;

            loop.Mode = dataOnly ? LoopMode.DataOnly : LoopMode.InsertRows;
            loop.RowCopyCount = rowCopyCount;
            loop.Name = args[1];
            var enumerable = (await converter.GetData(args[0].Substring(1)))?.Value as IEnumerable;
            if (enumerable == null) return false;
            loop.Items = enumerable.OfType<object?>().ToList();
            return true;
        }

        static bool ParsePagedLoop(string text, LoopPlan loop, string sheetName, IReadOnlyList<ExcelPageLoopProcessor.PageLoopPlan> pagePlans)
        {
            if (!text.StartsWith("#PagedLoopRows")) return false;
            var pageArgs = text.Replace("#PagedLoopRows", "").Replace("(", "").Replace(")", "").Split(',').Select(value => value.Trim()).ToArray();
            if (pageArgs.Length < 5 || !int.TryParse(pageArgs[4], out var blockRowCount)) return false;
            if (!Enum.TryParse<ExcelPageLoopProcessor.PageType>(pageArgs[0], out var pageType)) return false;

            var pagePlan = pagePlans.FirstOrDefault(item =>
                pageType == ExcelPageLoopProcessor.PageType.First && item.FirstPageSheetName == sheetName ||
                pageType == ExcelPageLoopProcessor.PageType.Body && item.BodyPageSheetNames.Contains(sheetName) ||
                pageType == ExcelPageLoopProcessor.PageType.Last && item.LastPageSheetName == sheetName);
            if (pagePlan == null) return false;

            loop.Mode = LoopMode.Paged;
            loop.Name = pageArgs[3];
            loop.RowCopyCount = blockRowCount;
            loop.Items = pageType switch
            {
                ExcelPageLoopProcessor.PageType.First => pagePlan.FirstPageItems,
                ExcelPageLoopProcessor.PageType.Body => pagePlan.BodyPageItems[pagePlan.BodyPageSheetNames.IndexOf(sheetName)],
                _ => pagePlan.LastPageItems
            };
            return true;
        }

        /// <summary>
        /// Classifies a used cell once so later phases can dispatch by meaning instead of reparsing its text.
        /// </summary>
        static TemplateCellKind GetOperationKind(string text, string? symbol)
        {
            if (symbol != null) return TemplateCellKind.Symbol;
            if (!text.StartsWith("#")) return TemplateCellKind.Literal;
            return IsLoopDirective(text) ? TemplateCellKind.Directive : TemplateCellKind.Function;
        }

        /// <summary>
        /// Selects the formatting inherited by a nested loop.
        /// Data-only loops keep destination formatting, while insert loops copy the template.
        /// </summary>
        static RowFormattingMode GetChildFormattingMode(LoopMode loopMode, RowFormattingMode parentMode)
            => loopMode switch
            {
                LoopMode.DataOnly => RowFormattingMode.PreserveDestination,
                LoopMode.InsertRows => RowFormattingMode.CopyTemplate,
                _ => parentMode
            };

        #endregion

        #region Materialization

        /// <summary>
        /// Selects the fast matrix strategy only when all cells use the default style.
        /// InsertData cannot preserve arbitrary styles or cell types.
        /// </summary>
        static bool CanUsePlainValueBatch(IXLWorksheet sheet, TemplateSnapshot template, ExpansionPlan plan)
            => template.CanUsePlainValueBatch && plan.Merges.Count == 0 && sheet.RowHeight == XLWorkbook.DefaultRowHeight && sheet.Rows(1, template.RowCount).All(row => row.Height == sheet.RowHeight);

        /// <summary>
        /// Clears old content and reserves the final row count in one operation.
        /// Repeated insertions make ClosedXML repeatedly update the rows below them.
        /// </summary>
        static void ReserveWorksheetRows(IXLWorksheet sheet, int outputRowCount, int sourceRowCount, int colCount)
        {
            foreach (var merge in sheet.MergedRanges.ToList()) merge.Unmerge();
            if (outputRowCount > sourceRowCount) sheet.Row(1).InsertRowsAbove(outputRowCount - sourceRowCount);

            var rowsToClear = Math.Max(sourceRowCount, outputRowCount);
            if (rowsToClear > 0 && colCount > 0) sheet.Range(1, 1, rowsToClear, colCount).Clear(XLClearOptions.Contents);
            if (outputRowCount < sourceRowCount && colCount > 0) sheet.Range(outputRowCount + 1, 1, sourceRowCount, colCount).Clear(XLClearOptions.All);
        }

        /// <summary>
        /// Resolves symbols into one rectangular matrix because a single InsertData call is much cheaper for large reports.
        /// </summary>
        static async Task WriteSymbolValuesAsBatchAsync(IXLWorksheet sheet, IReadOnlyList<OutputRowPlan> rows, int colCount)
        {
            var values = new object?[rows.Count][];
            for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                var row = rows[rowIndex];
                var rowValues = new object?[colCount];
                foreach (var operation in row.Operations)
                {
                    if (operation.Kind != TemplateCellKind.Symbol) continue;
                    if (row.ShouldClearSymbols)
                    {
                        // Empty data-only rows must not retain the original symbol text.
                        rowValues[operation.ColumnNumber - 1] = null;
                        continue;
                    }
                    var symbol = operation.Symbol!;
                    if (row.ScopesAreSynchronous)
                    {
                        // Keep the built-in converter synchronous; this loop can process hundreds of thousands of cells.
                        rowValues[operation.ColumnNumber - 1] = TryResolveSymbolSynchronously(row, symbol, out var value)
                            ? value
                            : operation.TemplateText;
                    }
                    else
                    {
                        var result = await ResolveSymbolAsync(row.ConverterScopes, symbol);
                        rowValues[operation.ColumnNumber - 1] = result.Found ? result.Value : operation.TemplateText;
                    }
                }
                values[rowIndex] = rowValues;
            }

            // InsertData writes the ordinary cells in one ClosedXML operation.
            if (rows.Count > 0) sheet.Cell(1, 1).InsertData(values);
        }

        /// <summary>
        /// Copies formatted rows in groups and data-only rows as content.
        /// Data-only loops must preserve the styles already assigned to their destination rows.
        /// </summary>
        static void CopyTemplateRanges(IXLWorksheet sheet, IXLWorksheet source, IReadOnlyList<OutputRowPlan> rows, int sourceRowCount, int colCount)
        {
            var startIndex = 0;
            while (startIndex < rows.Count)
            {
                if (rows[startIndex].FormattingMode != RowFormattingMode.CopyTemplate)
                {
                    CopyValuesOnly(sheet, source, startIndex + 1, rows[startIndex].TemplateRow, sourceRowCount, colCount);
                    startIndex++;
                    continue;
                }

                var endExclusive = startIndex + 1;
                while (endExclusive < rows.Count
                    && rows[endExclusive].FormattingMode == RowFormattingMode.CopyTemplate
                    && rows[endExclusive].TemplateRow == rows[endExclusive - 1].TemplateRow + 1
                    && rows[endExclusive].TemplateRow <= sourceRowCount)
                    endExclusive++;
                var sourceStart = rows[startIndex].TemplateRow;
                var sourceEnd = rows[endExclusive - 1].TemplateRow;
                if (sourceStart > 0 && sourceEnd <= sourceRowCount)
                {
                    var destinationStart = startIndex + 1;
                    source.Range(sourceStart, 1, sourceEnd, colCount).CopyTo(sheet.Range(destinationStart, 1, destinationStart + sourceEnd - sourceStart, colCount));
                    for (var offset = 0; offset <= sourceEnd - sourceStart; offset++) CopyRowMetadata(sheet.Row(destinationStart + offset), source.Row(sourceStart + offset));
                }
                startIndex = endExclusive;
            }
        }

        /// <summary>
        /// Applies cached source styles to grouped destination ranges because style-only templates do not need full cell copies.
        /// </summary>
        static void ApplyStyleGroups(IXLWorksheet sheet, IXLWorksheet source, IReadOnlyList<OutputRowPlan> rows, int sourceRowCount, int colCount)
        {
            var styleRunsBySourceRow = new Dictionary<int, IReadOnlyList<StyleRun>>();
            var startIndex = 0;
            while (startIndex < rows.Count)
            {
                if (rows[startIndex].FormattingMode != RowFormattingMode.CopyTemplate)
                {
                    startIndex++;
                    continue;
                }

                var endExclusive = startIndex + 1;
                while (endExclusive < rows.Count
                    && rows[endExclusive].FormattingMode == RowFormattingMode.CopyTemplate
                    && rows[endExclusive].TemplateRow == rows[endExclusive - 1].TemplateRow)
                    endExclusive++;
                var sourceRow = rows[startIndex].TemplateRow;
                if (sourceRow > 0 && sourceRow <= sourceRowCount)
                {
                    if (!styleRunsBySourceRow.TryGetValue(sourceRow, out var styleRuns))
                    {
                        styleRuns = GetStyleRuns(source, sourceRow, colCount);
                        styleRunsBySourceRow[sourceRow] = styleRuns;
                    }

                    foreach (var styleRun in styleRuns)
                        sheet.Range(startIndex + 1, styleRun.FirstColumn, endExclusive, styleRun.LastColumn).Style = styleRun.Style;

                    var sourceMetadata = source.Row(sourceRow);
                    if (sourceMetadata.Height != XLWorkbook.DefaultRowHeight || sourceMetadata.IsHidden || sourceMetadata.OutlineLevel != 0)
                    {
                        for (var destinationRow = startIndex + 1; destinationRow <= endExclusive; destinationRow++)
                            CopyRowMetadata(sheet.Row(destinationRow), sourceMetadata);
                    }
                }

                startIndex = endExclusive;
            }
        }

        /// <summary>
        /// Groups adjacent columns with the same style so they can be formatted together.
        /// </summary>
        static IReadOnlyList<StyleRun> GetStyleRuns(IXLWorksheet source, int sourceRow, int colCount)
        {
            var segments = new List<StyleRun>();
            if (colCount == 0) return segments;

            var firstColumn = 1;
            var currentStyle = source.Cell(sourceRow, firstColumn).Style;
            for (var column = 2; column <= colCount; column++)
            {
                var style = source.Cell(sourceRow, column).Style;
                if (style.Equals(currentStyle)) continue;
                segments.Add(new StyleRun(firstColumn, column - 1, currentStyle));
                firstColumn = column;
                currentStyle = style;
            }

            segments.Add(new StyleRun(firstColumn, colCount, currentStyle));
            return segments;
        }

        /// <summary>
        /// Copies values and formulas without styles for data-only rows.
        /// Their preallocated destination cells own the styles.
        /// </summary>
        static void CopyValuesOnly(IXLWorksheet destination, IXLWorksheet source, int destinationRow, int sourceRow, int sourceRowCount, int colCount)
        {
            if (sourceRow <= 0 || sourceRow > sourceRowCount) return;
            for (var column = 1; column <= colCount; column++)
            {
                var sourceCell = source.Cell(sourceRow, column);
                var destinationCell = destination.Cell(destinationRow, column);
                if (sourceCell.HasFormula) destinationCell.FormulaA1 = sourceCell.FormulaA1;
                else destinationCell.Value = sourceCell.Value;
            }
        }

        /// <summary>
        /// Copies row height, visibility, and outline level because range copying does not reliably copy them.
        /// </summary>
        static void CopyRowMetadata(IXLRow destination, IXLRow source)
        {
            destination.Height = source.Height;
            if (source.IsHidden) destination.Hide(); else destination.Unhide();
            destination.OutlineLevel = source.OutlineLevel;
        }

        /// <summary>
        /// Recreates each planned merge because row expansion does not translate merge addresses.
        /// </summary>
        static void RestoreMergedRanges(IXLWorksheet sheet, IReadOnlyList<OutputRowPlan> rows, IReadOnlyList<MergeOperation> merges)
        {
            var restored = new HashSet<string>();
            foreach (var merge in merges)
            {
                var length = merge.LastRow - merge.FirstRow;
                for (var start = 0; start + length < rows.Count; start++)
                {
                    if (Enumerable.Range(0, length + 1).Any(offset => rows[start + offset].TemplateRow != merge.FirstRow + offset)) continue;
                    var destination = sheet.Range(start + 1, merge.FirstColumn, start + length + 1, merge.LastColumn);
                    if (restored.Add(destination.RangeAddress.ToString()!)) destination.Merge();
                }
            }
        }

        #endregion

        #region Cell operations

        /// <summary>
        /// Resolves symbols and runs functions after layout is final.
        /// Custom functions need the final row coordinates.
        /// </summary>
        static async Task ApplyCellOperationsAsync(IXLWorksheet sheet, IReadOnlyList<OutputRowPlan> rows)
        {
            for (var index = 0; index < rows.Count; index++)
            {
                var row = rows[index];
                foreach (var operation in row.Operations)
                {
                    if (row.ShouldClearSymbols && operation.Kind == TemplateCellKind.Symbol)
                        SetValue(sheet, index + 1, operation.ColumnNumber, null);

                    if (row.ShouldClearDirective && operation.Kind == TemplateCellKind.Directive && operation.ColumnNumber == 1)
                    {
                        // Range copying also copies the marker, so clear it explicitly.
                        SetValue(sheet, index + 1, operation.ColumnNumber, null);
                        continue;
                    }

                    switch (operation.Kind)
                    {
                        case TemplateCellKind.Function:
                        {
                            foreach (var function in Functions)
                            {
                                if (!operation.TemplateText.StartsWith($"#{function.Name}(", StringComparison.Ordinal)) continue;
                                var argsText = operation.TemplateText.Replace($"#{function.Name}", "").Replace("(", "").Replace(")", "");
                                var args = new List<object?>();
                                foreach (var argument in argsText.Split(',').Select(value => value.Trim()))
                                    args.Add(argument.StartsWith("$") ? (await ResolveSymbolAsync(row.ConverterScopes, argument.Substring(1))).Value : argument);
                                await function.InvokeAsync(sheet, index + 1, operation.ColumnNumber, args.ToArray());
                                break;
                            }
                            break;
                        }
                        case TemplateCellKind.Symbol:
                        {
                            var result = await ResolveSymbolAsync(row.ConverterScopes, operation.Symbol!);
                            if (result.Found) SetValue(sheet, index + 1, operation.ColumnNumber, result.Value);
                            break;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Searches inner scopes first, then outer scopes.
        /// This lets nested values override root values while keeping outer references available.
        /// </summary>
        static async Task<(bool Found, object? Value)> ResolveSymbolAsync(IReadOnlyList<IExcelSymbolConverter> converterScopes, string symbol)
        {
            for (var index = converterScopes.Count - 1; index >= 0; index--)
            {
                if (converterScopes[index] is ISynchronousExcelSymbolConverter synchronous)
                {
                    // Avoid creating a task for synchronous lookups.
                    if (synchronous.TryGetData(symbol, out var value)) return (true, value);
                }
                else
                {
                    var value = await converterScopes[index].GetData(symbol);
                    if (value != null) return (true, value.Value);
                }
            }
            return (false, null);
        }

        /// <summary>
        /// Resolves synchronous scopes without tasks to keep the hot path fast.
        /// </summary>
        static bool TryResolveSymbolSynchronously(OutputRowPlan row, string symbol, out object? value)
        {
            if (row.InnermostSynchronousConverter != null && row.InnermostSynchronousConverter.TryGetData(symbol, out value)) return true;
            for (var index = row.ConverterScopes.Count - 2; index >= 0; index--)
            {
                if (((ISynchronousExcelSymbolConverter)row.ConverterScopes[index]).TryGetData(symbol, out value)) return true;
            }

            value = null;
            return false;
        }

        /// <summary>
        /// Writes a CLR value through ClosedXML; symbols may produce strings, numbers, dates, or null.
        /// </summary>
        static void SetValue(IXLWorksheet sheet, int row, int column, object? value)
            => sheet.Cell(row, column).SetValue(XLCellValue.FromObject(value));

        /// <summary>
        /// Checks whether the text starts with a supported loop directive.
        /// </summary>
        static bool IsLoopDirective(string text)
            => text.StartsWith("#LoopRow(") || text.StartsWith("#LoopRowData(") || text.StartsWith("#PagedLoopRows(");

        #endregion

        #region Plan models

        /// <summary>
        /// Controls whether a planned row removes the directive or its unresolved symbols.
        /// A flag keeps the two independent cleanup actions explicit without adding per-row objects.
        /// </summary>
        [Flags]
        enum RowCleanup
        {
            None = 0,
            Directive = 1,
            Symbols = 2
        }

        /// <summary>
        /// Describes whether a row receives template formatting or keeps its destination formatting.
        /// This is the key difference between normal loops and data-only loops.
        /// </summary>
        enum RowFormattingMode
        {
            CopyTemplate,
            PreserveDestination
        }

        /// <summary>
        /// Identifies the structural behavior of a loop directive.
        /// </summary>
        enum LoopMode
        {
            InsertRows,
            DataOnly,
            Paged
        }

        /// <summary>
        /// Identifies the small set of cell operations the writer must revisit after copying the layout.
        /// Literal and formula cells stay in the copied template and need no per-cell operation here.
        /// </summary>
        enum TemplateCellKind
        {
            Literal,
            Symbol,
            Directive,
            Function
        }

        /// <summary>
        /// Stores source operations and row markers so they remain available after the worksheet changes.
        /// </summary>
        sealed class TemplateSnapshot
        {
            readonly IReadOnlyList<TemplateCellOperation>[] _rows;
            readonly string[] _leftText;

            public TemplateSnapshot(IReadOnlyList<TemplateCellOperation>[] rows, string[] leftText, bool canUsePlainValueBatch, bool canUseStyledValueBatch)
            {
                _rows = rows;
                _leftText = leftText;
                CanUsePlainValueBatch = canUsePlainValueBatch;
                CanUseStyledValueBatch = canUseStyledValueBatch;
            }

            public bool CanUsePlainValueBatch { get; }
            public bool CanUseStyledValueBatch { get; }
            public int RowCount => _rows.Length - 1;
            public IReadOnlyList<TemplateCellOperation> GetRow(int row) => row > 0 && row < _rows.Length ? _rows[row] : Array.Empty<TemplateCellOperation>();
            public string GetLeftText(int row) => row > 0 && row < _leftText.Length ? _leftText[row] : string.Empty;
        }

        /// <summary>
        /// Stores one parsed source cell so it does not need to be inspected for every generated row.
        /// </summary>
        sealed class TemplateCellOperation
        {
            public TemplateCellOperation(int columnNumber, string templateText, string? symbol, TemplateCellKind kind)
            {
                ColumnNumber = columnNumber;
                TemplateText = templateText;
                Symbol = symbol;
                Kind = kind;
            }

            public int ColumnNumber { get; }
            public string TemplateText { get; }
            public string? Symbol { get; }
            public TemplateCellKind Kind { get; }
        }

        /// <summary>
        /// Stores one output row, its converter scopes, and the cleanup rules inherited from its loop.
        /// </summary>
        sealed class OutputRowPlan
        {
            public OutputRowPlan(int templateRow, IReadOnlyList<IExcelSymbolConverter> converterScopes, RowCleanup cleanup, RowFormattingMode formattingMode, IReadOnlyList<TemplateCellOperation> operations)
            {
                TemplateRow = templateRow;
                ConverterScopes = converterScopes;
                ScopesAreSynchronous = converterScopes.All(scope => scope is ISynchronousExcelSymbolConverter);
                InnermostSynchronousConverter = converterScopes[^1] as ISynchronousExcelSymbolConverter;
                Cleanup = cleanup;
                FormattingMode = formattingMode;
                Operations = operations;
            }

            public int TemplateRow { get; }
            public IReadOnlyList<IExcelSymbolConverter> ConverterScopes { get; }
            public bool ScopesAreSynchronous { get; }
            public ISynchronousExcelSymbolConverter? InnermostSynchronousConverter { get; }
            public RowCleanup Cleanup { get; }
            public RowFormattingMode FormattingMode { get; }
            public IReadOnlyList<TemplateCellOperation> Operations { get; }
            public bool ShouldClearDirective => (Cleanup & RowCleanup.Directive) != 0;
            public bool ShouldClearSymbols => (Cleanup & RowCleanup.Symbols) != 0;
        }

        /// <summary>
        /// Stores one source merge so each repeated occurrence can get translated coordinates.
        /// </summary>
        readonly record struct MergeOperation(int FirstRow, int LastRow, int FirstColumn, int LastColumn);

        readonly record struct AutoFilterSnapshot(int FirstRow, int FirstColumn, int LastRow, int LastColumn);

        /// <summary>
        /// Describes a contiguous style run because repeated destination rows can receive it as one range.
        /// </summary>
        readonly record struct StyleRun(int FirstColumn, int LastColumn, IXLStyle Style);

        /// <summary>
        /// Stores the final row layout so materialization does not rediscover loop structure.
        /// </summary>
        sealed class ExpansionPlan
        {
            public ExpansionPlan(IReadOnlyList<OutputRowPlan> rows, IReadOnlyList<MergeOperation> merges)
            {
                Rows = rows;
                Merges = merges;
            }

            public IReadOnlyList<OutputRowPlan> Rows { get; }
            public IReadOnlyList<MergeOperation> Merges { get; }
        }

        /// <summary>
        /// Stores the settings shared by normal, data-only, and paged loops.
        /// </summary>
        sealed class LoopPlan
        {
            public int RowCopyCount { get; set; }
            public List<object?> Items { get; set; } = new();
            public string Name { get; set; } = string.Empty;
            public LoopMode Mode { get; set; }
        }

        #endregion
    }
}
