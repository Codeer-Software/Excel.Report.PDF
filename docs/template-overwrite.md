# Template overwrite

`Excel.Report.PDF` ships with a small but expressive template engine. You author a regular `.xlsx` file in Excel, place `$symbols` and `#directives` in the cells you want to bind, and at runtime the library overwrites those cells with values from your data object.

## TL;DR

```csharp
using ClosedXML.Excel;
using Excel.Report.PDF;

using var book = new XLWorkbook("template.xlsx");
await book.Worksheet(1).OverWrite(new ObjectExcelSymbolConverter(myData));
book.SaveAs("populated.xlsx");
```

To overwrite **every** sheet (and trigger the multi-page expansion described in [multi-page.md](multi-page.md)):

```csharp
await book.OverWrite(new ObjectExcelSymbolConverter(myData));
```

## The two extension methods

Both are defined in `Excel.Report.PDF.ExcelOverWriter`:

| Method | Scope |
| --- | --- |
| `XLWorkbook.OverWrite(IExcelSymbolConverter)` | All sheets. Required when the workbook contains `#PagedLoopRows` because that directive may add or replace sheets. |
| `IXLWorksheet.OverWrite(IExcelSymbolConverter)` | A single sheet only. |

Both methods are `async Task`. Awaiting is required if your `IExcelSymbolConverter` performs I/O.

## Symbols (`$`)

A cell whose **trimmed** text starts with `$` is treated as a symbol. The text after `$` is passed to `IExcelSymbolConverter.GetData(string)`. If the converter returns a non-null `ExcelOverWriteCell`, the cell is overwritten with `ExcelOverWriteCell.Value`.

```text
| Title:   | $Client                |
| Charge:  | $PersonInCharge        |
```

With `ObjectExcelSymbolConverter(data)` and `data.Client = "Acme"`, the cell is overwritten with `Acme`.

`ObjectExcelSymbolConverter` resolves symbols against the bound object's public properties using simple reflection (`GetType().GetProperty(symbol)`). Returning `null` from `GetData` leaves the cell untouched.

### Loop-scoped symbols

When iterating over a list (see below), symbols are evaluated against the current loop element by prefixing the property path with `loopName.`:

```text
#LoopRow($Details, item, 1)
| $item.Title | $item.Detail | $item.Price | $item.Discount |
```

This matches `ObjectExcelSymbolConverter.GetData(object element, string elementName, string symbol)`, which strips the `elementName.` prefix and resolves the rest against the element's properties.

## Loop directives

Loop directives must live in **column A** because the engine scans column A row-by-row to detect them. The directive cell is cleared after expansion, so column A is free to keep your control instructions invisible in the final document.

### `#LoopRow($items, itemName, rowCopyCount)`

Insert mode. The engine **inserts** `rowCopyCount` rows above the current row for each element in `$items`, copies the row format, and writes the values.

* `$items` — a property whose value implements `IEnumerable`. The leading `$` is required.
* `itemName` — the name used in `$itemName.Property` placeholders inside the loop block.
* `rowCopyCount` — optional. How many rows form a single block. Defaults to `1`. Use this when each record spans multiple rows in the template.

When `$items` is empty, the loop block (the `rowCopyCount` rows starting at the directive row) is **deleted** so the surrounding layout collapses.

### `#LoopRowData($items, itemName, rowCopyCount)`

Data-only mode. The original row format is reused without inserting rows. Use this when the template already contains pre-formatted rows (for example, a fixed-size grid) and you want to fill them in place.

When `$items` is empty in data mode, the cells inside the block whose text starts with `$` are cleared, and other cells are left intact.

### Nested loops

Loops can be nested. The engine recursively re-runs `OverWrite` on the freshly copied block, with a **child symbol converter** scoped to the current element:

```csharp
public interface IExcelSymbolConverter
{
    IExcelSymbolConverter CreateChildExcelSymbolConverter(object? obj, string name);
    Task<ExcelOverWriteCell?> GetData(string symbol);
}
```

`ObjectExcelSymbolConverter.CreateChildExcelSymbolConverter(element, name)` returns a converter that resolves `name.Property` against `element`. The outer converter still resolves top-level symbols, so nested loops can reference outer values too — see the test fixtures `RecursiveLoop1Test` and `RecursiveLoop2Test` for end-to-end examples.

## `#PagedLoopRows`

`#PagedLoopRows` is the multi-page sibling of `#LoopRow`. It distributes one logical list across three template sheets named **First**, **Body**, and **Last**, dynamically duplicating the body sheet as many times as needed.

It is documented separately in [multi-page.md](multi-page.md).

## Implementing your own `IExcelSymbolConverter`

`ObjectExcelSymbolConverter` is a small, reflection-based default. For non-POCO data sources (database rows, JSON, locale-aware formatting, async lookups, etc.) implement the interface yourself.

```csharp
public sealed class DictionaryConverter : IExcelSymbolConverter
{
    private readonly IReadOnlyDictionary<string, object?> _root;
    private readonly object? _scope;
    private readonly string _scopeName;

    public DictionaryConverter(IReadOnlyDictionary<string, object?> root)
        : this(root, null, string.Empty) { }

    private DictionaryConverter(IReadOnlyDictionary<string, object?> root, object? scope, string scopeName)
    {
        _root = root;
        _scope = scope;
        _scopeName = scopeName;
    }

    public IExcelSymbolConverter CreateChildExcelSymbolConverter(object? obj, string name)
        => new DictionaryConverter(_root, obj, name);

    public Task<ExcelOverWriteCell?> GetData(string symbol)
    {
        // Nested loop scope — symbol arrives as "elementName.Property".
        if (!string.IsNullOrEmpty(_scopeName) && symbol.StartsWith(_scopeName + "."))
        {
            var prop = symbol.Substring(_scopeName.Length + 1);
            var value = _scope?.GetType().GetProperty(prop)?.GetValue(_scope);
            return Task.FromResult<ExcelOverWriteCell?>(new() { Value = value });
        }

        // Top-level lookup against the dictionary.
        return Task.FromResult<ExcelOverWriteCell?>(
            _root.TryGetValue(symbol, out var v) ? new() { Value = v } : null);
    }
}
```

Things to honour when writing a custom converter:

1. **Return `null`** for unknown symbols. The engine treats `null` as "leave the cell alone" and only `ExcelOverWriteCell.Value` writes through.
2. **Provide a child scope** for loops via `CreateChildExcelSymbolConverter`. Without it, `$item.Property` placeholders cannot be resolved inside loop bodies.
3. **`IEnumerable`** is required for loop sources. Anything that implements `IEnumerable` (arrays, `List<T>`, LINQ queries, `IAsyncEnumerable` materialised into a list, etc.) works.

## Best practices

* Author the template in Excel, save as `.xlsx`, and version-control it next to the code that binds it.
* Keep loop directive cells in column A. Use a hidden column width or a white font if you do not want them visible while editing.
* When a directive does not fire, the most common cause is a typo (the engine matches `#LoopRow` / `#LoopRowData` / `#PagedLoopRows` exactly) or a stray apostrophe that turns the cell into text without `$` recognition.
* For numbers and dates, format the destination cell in Excel — the value written by `Value = ...` flows through `XLCellValue.FromObject`, so types are preserved and the cell's number format applies.

## See also

* Worked sample workbooks live under `Source/Test/Test/PdfSrc/` (used by the unit tests in `ExcelOverWriterTest.cs`).
* [multi-page.md](multi-page.md) — the multi-page form of looping.
* [built-in-functions.md](built-in-functions.md) — `#Image` and `#QR`, plus how to register your own `#MyFunction`.
