# Public API reference

This page is a compact reference for every public type and member exposed by `Excel.Report.PDF` and `Excel.Report.PrintDocument`. For richer narrative, follow the cross-links to the topic guides.

## Namespace `Excel.Report.PDF`

### `static class ExcelConverter`

Entry point for Excel → PDF conversion. All overloads return a `MemoryStream` containing the produced PDF bytes; the caller owns the stream.

| Method | Description |
| --- | --- |
| `MemoryStream ConvertToPdf(string filePath)` | Convert every sheet in the workbook at `filePath`. |
| `MemoryStream ConvertToPdf(Stream stream)` | Convert every sheet from a workbook stream. |
| `MemoryStream ConvertToPdf(string filePath, int sheetPosition)` | Convert a single 1-based sheet position. |
| `MemoryStream ConvertToPdf(Stream stream, int sheetPosition)` | Stream variant of the above. |
| `MemoryStream ConvertToPdf(string filePath, string sheetName)` | Convert a single sheet by name. |
| `MemoryStream ConvertToPdf(Stream stream, string sheetName)` | Stream variant of the above. |

See [getting-started.md](getting-started.md).

### `static class ExcelOverWriter`

Extension methods on ClosedXML types that resolve `$symbols` and `#directives` against an `IExcelSymbolConverter`. Both methods are `async`.

| Method | Description |
| --- | --- |
| `Task XLWorkbook.OverWrite(IExcelSymbolConverter converter)` | Overwrite every sheet, expanding `#PagedLoopRows` body templates as needed. |
| `Task IXLWorksheet.OverWrite(IExcelSymbolConverter converter)` | Overwrite a single sheet only (no multi-page expansion). |
| `void RegisterOverWriteFunction(IOverWriteFunction function)` | Register a custom `#FunctionName(...)` directive globally. |

See [template-overwrite.md](template-overwrite.md) and [multi-page.md](multi-page.md).

### `interface IExcelSymbolConverter`

```csharp
public interface IExcelSymbolConverter
{
    IExcelSymbolConverter CreateChildExcelSymbolConverter(object? obj, string name);
    Task<ExcelOverWriteCell?> GetData(string symbol);
}
```

* `GetData(symbol)` resolves a top-level `$symbol` (or a `loopName.Property` symbol when in a child scope). Return `null` to leave the cell unchanged.
* `CreateChildExcelSymbolConverter(element, name)` is called once per loop element. The returned converter is used while the engine recursively re-runs `OverWrite` on the duplicated block.

### `class ObjectExcelSymbolConverter : IExcelSymbolConverter`

Reflection-based default. Resolves symbols against the public properties of the object passed to the constructor. Loop scopes resolve `loopName.Property` against the current element.

```csharp
new ObjectExcelSymbolConverter(myDataObject)
```

### `class ExcelOverWriteCell`

Carries the value to be written:

```csharp
public class ExcelOverWriteCell
{
    public object? Value { get; set; }
}
```

The value is funnelled through `XLCellValue.FromObject(...)`, so primitives, strings, dates, and `null` work out of the box.

### `interface IOverWriteFunction`

Implement to register a custom `#Name(args...)` directive.

```csharp
public interface IOverWriteFunction
{
    string Name { get; }
    Task InvokeAsync(IXLWorksheet sheet, int rowIndex, int colIndex, object?[] args);
}
```

The engine matches `#{Name}(` exactly. `args` already has `$` symbols resolved.

See [built-in-functions.md](built-in-functions.md).

### `interface IPostProcessCommand` & `PostProcessCommandExtensions`

```csharp
public interface IPostProcessCommand { void Execute(); }

public static class PostProcessCommandExtensions
{
    public static void ExecuteAll(this IEnumerable<IPostProcessCommand> commands);
}
```

Used internally to back-fill `#PageCount` after pagination is complete. Public so advanced consumers can plug into the same mechanism if they extend the rendering pipeline.

### `static class ExcelUtils`

Convenience helpers for working with worksheet contents.

| Method | Description |
| --- | --- |
| `List<List<string>> ReadAllTexts(this IXLWorksheet sheet)` | Snapshot every used cell as trimmed text. |
| `Task<List<List<string>>> ReadAllTextsFromExcelBinary(Stream excel)` | Open a workbook from a stream and read sheet 1's text. |
| `List<List<object?>> ReadAllObjects(this IXLWorksheet sheet)` | Same shape as `ReadAllTexts` but typed by `XLDataType`. |
| `MemoryStream CreateExcelBinary(List<List<string>> allTexts, string sheetName)` | Build a workbook from a 2-D string grid. |
| `MemoryStream CreateExcelBinary(List<List<object?>> objects, string sheetName)` | Same as above but for typed values. |

## Namespace `Excel.Report.PrintDocument`

> Windows-only — annotated `[SupportedOSPlatform("windows")]`.

### `class ExcelPrintDocumentBinder`

| Method | Description |
| --- | --- |
| `static void Bind(System.Drawing.Printing.PrintDocument doc, string filePath, PrintPageSetup? setup = null)` | Wire a workbook on disk into a `PrintDocument`. |
| `static void Bind(System.Drawing.Printing.PrintDocument doc, Stream stream, PrintPageSetup? setup = null)` | Stream variant. |
| `static PrintPageSetup GetPrintPageSetup(string filePath)` | Read the page setup from the first worksheet. |
| `static PrintPageSetup GetPrintPageSetup(Stream stream)` | Stream variant. |

`Bind` subscribes to `PrintPage` and `EndPrint` and detaches both handlers automatically when printing finishes, so the same `PrintDocument` instance can be reused.

See [print-document.md](print-document.md).

### `class PrintPageSetup`

```csharp
public class PrintPageSetup
{
    public PrintMargins Margins { get; set; } = new PrintMargins();
    public double WidthPoint  { get; set; }
    public double HeightPoint { get; set; }

    public PrintPageSetup FromIXLPageSetup(IXLPageSetup pageSetup);
    public static double MmToPoint(double mm);
}
```

All measurements are in PostScript points (1 pt = 1/72 inch).

* `FromIXLPageSetup` initialises the instance from a ClosedXML `IXLPageSetup`.
* `MmToPoint` is provided as a utility for callers who think in millimetres.

### `class PrintMargins`

```csharp
public class PrintMargins
{
    public double LeftPoint   { get; set; }
    public double RightPoint  { get; set; }
    public double TopPoint    { get; set; }
    public double BottomPoint { get; set; }
}
```

## Cell directive cheat sheet

| Directive | Allowed location | Detail |
| --- | --- | --- |
| `$name` | any cell | [template-overwrite.md](template-overwrite.md) |
| `#LoopRow($items, item, n)` | column **A** | [template-overwrite.md](template-overwrite.md#looprowitems-itemname-rowcopycount) |
| `#LoopRowData($items, item, n)` | column **A** | [template-overwrite.md](template-overwrite.md#looprowdataitems-itemname-rowcopycount) |
| `#PagedLoopRows(pageType, rowsPerPage, $items, item, blockRowCount)` | column **A** | [multi-page.md](multi-page.md) |
| `#Image($bytesOrStream[, w[, h]])` | any cell | [built-in-functions.md](built-in-functions.md#image) |
| `#QR($text[, pixelsPerModule])` | any cell | [built-in-functions.md](built-in-functions.md#qr) |
| `#Page` | any cell except column A | [special-directives.md](special-directives.md#page-number-directives) |
| `#PageCount` | any cell except column A | [special-directives.md](special-directives.md#page-number-directives) |
| `#PageOf("/")` | any cell except column A | [special-directives.md](special-directives.md#page-number-directives) |
| `#Empty` | any cell | [special-directives.md](special-directives.md#empty) |
| `#FitColumn` | **A1** only | [special-directives.md](special-directives.md#fitcolumn) |
