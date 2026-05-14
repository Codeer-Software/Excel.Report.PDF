# Cell functions & extensibility

Cell functions are directives that start with `#FunctionName(...)`. They are invoked during workbook overwrite (after `$` symbols are resolved) and can rewrite the cell, insert pictures, or run any custom logic you can express in C#.

The engine has **no hard-coded functions**. Even the two built-ins ship as ordinary `IOverWriteFunction` implementations that you can study, replace, or augment.

| Function | Source | Purpose |
| --- | --- | --- |
| [`#Image`](#image) | [`ImageOverWriteFunction.cs`](../Source/Excel.Report.PDF/ImageOverWriteFunction.cs) | Insert a picture from `byte[]` or `Stream`. |
| [`#QR`](#qr) | [`QRCodeOverWriteFunction.cs`](../Source/Excel.Report.PDF/QRCodeOverWriteFunction.cs) | Generate and insert a QR code. |

Jump to [Writing your own function](#writing-your-own-function) for the extension story.

---

## `#Image`

```text
#Image($bytesOrStream)
#Image($bytesOrStream, widthScale)
#Image($bytesOrStream, widthScale, heightScale)
```

| Argument | Type | Required | Description |
| --- | --- | --- | --- |
| `$bytesOrStream` | `byte[]` or `Stream` | yes | Source data for the image. The leading `$` resolves through your `IExcelSymbolConverter`. |
| `widthScale` | `double` | optional | Multiplied with the image's natural width via `IXLPicture.ScaleWidth`. |
| `heightScale` | `double` | optional | Multiplied with the image's natural height via `IXLPicture.ScaleHeight`. |

Behaviour:

* The image is anchored to the directive cell (top-left, no offset) using ClosedXML's `IXLPicture.MoveTo(...)`.
* The directive cell text is cleared.
* If `$bytesOrStream` resolves to `null` or to a non-stream/non-`byte[]` value, the call is a no-op and the cell is left as-is.

```csharp
public class Item
{
    public byte[] Logo { get; set; } = Array.Empty<byte>();
}
```

```text
| #Image($Logo, 0.5, 0.5) |
```

## `#QR`

```text
#QR($text)
#QR($text, pixelsPerModule)
```

| Argument | Type | Required | Description |
| --- | --- | --- | --- |
| `$text` or `"literal"` | `string` | yes | The text encoded in the QR code. Quoted literals (e.g. `"https://example.com"`) are also accepted; the surrounding quotes are stripped. |
| `pixelsPerModule` | `int` | optional | Module size in pixels (default `10`). Larger values produce a larger image. |

Properties:

* Generated as PNG with **error correction level M**.
* Quiet zones are drawn (`drawQuietZones: true`).
* Anchored to the directive cell (top-left, no offset). The cell text is cleared.

```csharp
public class Card
{
    public string Url { get; set; } = "https://www.codeer.co.jp/";
}
```

```text
| #QR($Url, 8) |
```

---

## Writing your own function

The whole extensibility story is a single, two-method interface and one registration call.

### The interface

```csharp
namespace Excel.Report.PDF;

public interface IOverWriteFunction
{
    string Name { get; }   // matched as "#{Name}("

    Task InvokeAsync(
        IXLWorksheet sheet,
        int rowIndex,
        int colIndex,
        object?[] args);
}
```

| Member | Notes |
| --- | --- |
| `Name` | The text that follows `#`. The engine matches `#{Name}(` **exactly** (case-sensitive), so keep it short, distinctive, and avoid colliding with `LoopRow`, `LoopRowData`, `PagedLoopRows`, `Page`, `PageCount`, `PageOf`, `Empty`, `FitColumn`, `Image`, `QR`. |
| `InvokeAsync` | Called once per matching cell. `args` is already populated with the resolved arguments (see below). The function decides how to mutate the cell (or any other cell on the sheet). Return a completed `Task` if the work is synchronous. |

### Registering

```csharp
using Excel.Report.PDF;

ExcelOverWriter.RegisterOverWriteFunction(new MyFunction());
```

Notes on lifecycle:

* The registration list is **process-wide** — register once at application startup, before the first `OverWrite` call.
* Re-registering the same instance appends a duplicate, which is harmless but pointless. Guard with a flag if your startup runs multiple times (tests, hot reload, etc.).
* Built-in functions (`#Image`, `#QR`) are registered automatically. To disable them, simply do not invoke them from your templates — the registration cannot be removed.

### How `args` is built

When the engine sees `#MyFunc(arg1, arg2, arg3)` in a cell, it:

1. Strips `#MyFunc(` and the trailing `)`.
2. Splits the remainder on `,` and trims each piece.
3. For each piece:
   * Starts with `$` → calls `IExcelSymbolConverter.GetData(name)` and stores the resolved `ExcelOverWriteCell.Value` (which can be `null`).
   * Otherwise → passes the trimmed `string` through unchanged.
4. Hands the resulting `object?[]` to `InvokeAsync`.

So `#QR($Url, 8)` with `Url = "https://example.com"` arrives as `["https://example.com", "8"]`. It is the function's responsibility to parse `"8"` into the type it actually wants:

```csharp
int size = 10;
if (1 < args.Length && int.TryParse(args[1]?.ToString(), out var v)) size = v;
```

> Heads-up: arguments are split on `,` *literally*. If an argument may contain a comma (e.g. CSV text or a long URL with query parameters), bind it via `$symbol` so the comma never appears in the directive text.

### Recipes

Below are short, self-contained examples covering common needs. Each one is a complete `IOverWriteFunction` implementation.

#### 1. Format a value with culture-aware currency

```csharp
public sealed class YenFunction : IOverWriteFunction
{
    public string Name => "Yen";

    public Task InvokeAsync(IXLWorksheet sheet, int rowIndex, int colIndex, object?[] args)
    {
        var amount = Convert.ToDecimal(args.ElementAtOrDefault(0) ?? 0m);
        var formatted = amount.ToString("¥#,##0", CultureInfo.GetCultureInfo("ja-JP"));
        sheet.Cell(rowIndex, colIndex).SetValue(XLCellValue.FromObject(formatted));
        return Task.CompletedTask;
    }
}
```

Use as `#Yen($Total)`.

#### 2. Stamp a signature image conditionally

```csharp
public sealed class SignatureFunction : IOverWriteFunction
{
    public string Name => "Signature";

    public Task InvokeAsync(IXLWorksheet sheet, int rowIndex, int colIndex, object?[] args)
    {
        var bytes = args.ElementAtOrDefault(0) as byte[];
        sheet.Cell(rowIndex, colIndex).SetValue(XLCellValue.FromObject(null));

        if (bytes is null || bytes.Length == 0)
            return Task.CompletedTask;     // approver not yet signed → leave blank

        using var stream = new MemoryStream(bytes);
        var picture = sheet.AddPicture(stream).MoveTo(sheet.Cell(rowIndex, colIndex));
        picture.ScaleWidth(0.6);
        picture.ScaleHeight(0.6);
        return Task.CompletedTask;
    }
}
```

Use as `#Signature($Approver.SignatureBytes)`.

#### 3. Async lookup against an external system

```csharp
public sealed class ProductLookupFunction : IOverWriteFunction
{
    private readonly IProductRepository _products;
    public ProductLookupFunction(IProductRepository products) => _products = products;

    public string Name => "Product";

    public async Task InvokeAsync(IXLWorksheet sheet, int rowIndex, int colIndex, object?[] args)
    {
        var sku = args.ElementAtOrDefault(0)?.ToString();
        if (string.IsNullOrEmpty(sku)) return;

        var product = await _products.FindBySkuAsync(sku);
        sheet.Cell(rowIndex, colIndex).SetValue(XLCellValue.FromObject(product?.DisplayName ?? "(unknown)"));
    }
}
```

Use as `#Product($Item.Sku)`.

#### 4. Render a 1-D barcode into the cell

```csharp
public sealed class BarcodeFunction : IOverWriteFunction
{
    public string Name => "Barcode";

    public Task InvokeAsync(IXLWorksheet sheet, int rowIndex, int colIndex, object?[] args)
    {
        var text = args.ElementAtOrDefault(0)?.ToString() ?? string.Empty;
        if (text.StartsWith("\"") && text.EndsWith("\"")) text = text[1..^1];
        if (string.IsNullOrEmpty(text)) return Task.CompletedTask;

        // Bring your favourite barcode lib (e.g., ZXing.Net, BarcodeStandard).
        // Anything that yields a PNG byte[] works.
        byte[] png = MyBarcodeLib.RenderCode128(text);

        using var stream = new MemoryStream(png);
        sheet.AddPicture(stream).MoveTo(sheet.Cell(rowIndex, colIndex));
        sheet.Cell(rowIndex, colIndex).SetValue(XLCellValue.FromObject(null));
        return Task.CompletedTask;
    }
}
```

Use as `#Barcode($Item.Code)` — exactly the pattern `#QR` follows internally.

### Reference: how `#Image` and `#QR` are implemented

The two built-ins are intentionally small and worth reading before you write your own:

* **`#Image`** — [`ImageOverWriteFunction.cs`](../Source/Excel.Report.PDF/ImageOverWriteFunction.cs)
  Demonstrates handling either `Stream` or `byte[]` input, optional width/height scaling, and clearing the source cell after positioning.
* **`#QR`** — [`QRCodeOverWriteFunction.cs`](../Source/Excel.Report.PDF/QRCodeOverWriteFunction.cs)
  Demonstrates accepting a quoted literal *or* a `$symbol`, parsing a numeric optional argument, and inserting an image generated on the fly via QRCoder.

Both follow the same shape: validate input → produce a stream → `sheet.AddPicture(...).MoveTo(...)` → clear the directive cell.

### Best practices

1. **Always clear the directive cell** if you do not write a meaningful value to it. Otherwise the literal `#YourFunc(...)` text remains in the workbook.
2. **Treat `args` defensively** — entries may be `null` (unresolved symbol), strings (literals), or arbitrary CLR objects (resolved symbols). Use `as` checks or `Convert.*` helpers.
3. **Avoid throwing.** Failures inside `InvokeAsync` propagate out of `OverWrite` and abort the whole render. Decide consciously whether a missing piece of data should fail loudly or render as blank/`(unknown)`.
4. **Stay synchronous when you can.** Most cell-level transformations are CPU-only; returning `Task.CompletedTask` is faster than `async Task` with no `await`.
5. **Scope shared state carefully.** Functions are singletons inside `ExcelOverWriter`'s static list. Inject dependencies (`IProductRepository`, `IClock`, etc.) through the constructor so they are easy to mock in tests.

## See also

* [template-overwrite.md](template-overwrite.md) — how `$symbols` and the symbol converter feed the function arguments.
* [api-reference.md](api-reference.md) — full signatures for `IOverWriteFunction`, `IExcelSymbolConverter`, `ExcelOverWriteCell`.
