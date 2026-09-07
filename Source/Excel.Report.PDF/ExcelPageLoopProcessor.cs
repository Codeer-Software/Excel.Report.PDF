using System.Collections;
using ClosedXML.Excel;

namespace Excel.Report.PDF
{
    /// <summary>
    /// Builds page plans and creates body-page worksheets before OverWrite runs.
    /// Paging stays separate because it can add and remove worksheets.
    /// </summary>
    internal static class ExcelPageLoopProcessor
    {
        /// <summary>
        /// Identifies a worksheet's role in a paged report.
        /// </summary>
        internal enum PageType
        {
            First,
            Body,
            Last
        }

        /// <summary>
        /// Stores the items and sheet names needed to process one paged loop.
        /// </summary>
        internal sealed class PageLoopPlan
        {
            public List<object?> AllItems { get; } = new();
            public string FirstPageSheetName { get; set; } = string.Empty;
            public int FirstPageItemCount { get; set; }
            public string BodyTemplateSheetName { get; set; } = string.Empty;
            public List<string> BodyPageSheetNames { get; } = new();
            public int BodyPageItemCount { get; set; }
            public string LastPageSheetName { get; set; } = string.Empty;
            public int LastPageItemCount { get; set; }
            public List<object?> FirstPageItems { get; set; } = new();
            public List<List<object?>> BodyPageItems { get; } = new();
            public List<object?> LastPageItems { get; set; } = new();
        }

        /// <summary>
        /// Copies the body template for each page and removes the original source sheet.
        /// </summary>
        public static void MaterializeBodyPageSheets(XLWorkbook book, Dictionary<string, PageLoopPlan> pagePlans)
        {
            foreach (var entry in pagePlans)
            {
                if (string.IsNullOrEmpty(entry.Value.BodyTemplateSheetName)) continue;
                var bodySheet = book.Worksheet(entry.Value.BodyTemplateSheetName);
                for (var i = 0; i < entry.Value.BodyPageItems.Count; i++)
                    bodySheet.CopyTo($"{entry.Value.BodyTemplateSheetName}_{i}", bodySheet.Position + i);

                book.Worksheets.Delete(entry.Value.BodyTemplateSheetName);
            }
        }

        /// <summary>
        /// Reads and splits paged-loop data once so converter work is not repeated for every page.
        /// </summary>
        public static async Task<Dictionary<string, PageLoopPlan>> BuildPagePlansAsync(XLWorkbook book, IExcelSymbolConverter converter)
        {
            var pagePlans = new Dictionary<string, PageLoopPlan>();
            foreach (var sheet in book.Worksheets)
            {
                ExcelUtils.GetRowColCount(sheet, out var rowCount, out _);
                var leftCells = new List<string>();
                var pagedLoopCount = 0;
                for (var i = 0; i <= rowCount; i++)
                {
                    var text = sheet.GetText(i + 1, 1).Trim();
                    if (text.StartsWith("#PagedLoopRows")) pagedLoopCount++;
                    if (pagedLoopCount > 1) throw new Exception($"One sheet can have only one #PagedLoopRows. SheetName:{sheet.Name}");
                    leftCells.Add(text);
                }

                foreach (var leftCell in leftCells)
                {
                    if (!leftCell.StartsWith("#PagedLoopRows")) continue;
                    var args = leftCell.Replace("#PagedLoopRows", "").Replace("(", "").Replace(")", "").Split(',').Select(e => e.Trim()).ToArray();
                    if (args.Length != 5 || !args[2].StartsWith("$")) break;

                    var items = args[2].Substring(1);
                    if (!pagePlans.TryGetValue(items, out var plan))
                    {
                        plan = new PageLoopPlan();
                        var enumerable = (await converter.GetData(items))?.Value as IEnumerable;
                        if (enumerable == null) break;
                        plan.AllItems.AddRange(enumerable.Cast<object?>());
                        pagePlans[items] = plan;
                    }

                    if (!Enum.TryParse<PageType>(args[0], out var pageType) ||
                        !int.TryParse(args[1], out var rowsPerPage) ||
                        !int.TryParse(args[4], out _)) break;

                    switch (pageType)
                    {
                        case PageType.First:
                            plan.FirstPageSheetName = sheet.Name;
                            plan.FirstPageItemCount = rowsPerPage;
                            break;
                        case PageType.Body:
                            plan.BodyTemplateSheetName = sheet.Name;
                            plan.BodyPageItemCount = rowsPerPage;
                            break;
                        case PageType.Last:
                            plan.LastPageSheetName = sheet.Name;
                            plan.LastPageItemCount = rowsPerPage;
                            break;
                    }

                    break;
                }
            }

            // Split each source list once because all page sheets use the same boundaries.
            foreach (var entry in pagePlans)
            {
                var plan = entry.Value;
                if (!plan.AllItems.Any()) continue;

                var bodyCount = plan.AllItems.Count - plan.FirstPageItemCount - plan.LastPageItemCount;
                if (bodyCount == 0)
                {
                    var firstPageCount = plan.FirstPageItemCount;
                    var lastPageCount = plan.AllItems.Count - plan.FirstPageItemCount;
                    if (lastPageCount <= 0)
                    {
                        lastPageCount = 1;
                        firstPageCount = plan.AllItems.Count - 1;
                    }

                    plan.FirstPageItems = plan.AllItems.Take(firstPageCount).ToList();
                    plan.LastPageItems = plan.AllItems.Skip(firstPageCount).Take(lastPageCount).ToList();
                }
                else
                {
                    var rest = plan.AllItems.Skip(plan.FirstPageItemCount).ToList();
                    plan.FirstPageItems = plan.AllItems.Take(plan.FirstPageItemCount).ToList();
                    while (plan.LastPageItemCount < rest.Count)
                    {
                        plan.BodyPageItems.Add(rest.Take(plan.BodyPageItemCount).ToList());
                        rest = rest.Skip(plan.BodyPageItemCount).ToList();
                    }

                    plan.LastPageItems = rest;
                    for (var i = 0; i < plan.BodyPageItems.Count; i++) plan.BodyPageSheetNames.Add($"{plan.BodyTemplateSheetName}_{i}");
                }
            }

            return pagePlans;
        }
    }
}
