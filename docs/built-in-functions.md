# Built-in cell functions

Cell functions are directives that start with `#FunctionName(...)`. They are invoked during workbook overwrite (after `$` symbols are resolved) and can rewrite the cell, insert pictures, or run any custom logic.

Two functions ship in the box:

* [`#Image`](#image) — insert a picture from `byte[]` or `Stream`.
* [`#QR`](#qr) — generate and insert a QR code.

You can [register additional functions](#writing-your-own-function) by implementing `IOverWriteFunction`.

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

## How the engine parses arguments

A directive's arguments are split on `,` and trimmed. Each argument is then resolved as follows:

* `$something` → calls `IExcelSymbolConverter.GetData("something")` and uses `ExcelOverWriteCell.Value`.
* anything else → passed through as a literal `string`.

The function implementation is responsible for parsing literals (e.g. `int.TryParse` for sizes). This is why `#QR($Url, 8)` works without quoting `8` — the function converts it.

## Writing your own function

Implement `Excel.Report.PDF.IOverWriteFunction` and register it with `ExcelOverWriter.RegisterOverWriteFunction`:

```csharp
public sealed class HtmlEntityFunction : IOverWriteFunction
{
    public string Name => "Html";  // matched as "#Html(...)"

    public async Task InvokeAsync(IXLWorksheet sheet, int rowIndex, int colIndex, object?[] args)
    {
        await Task.CompletedTask;

        var raw = args.ElementAtOrDefault(0)?.ToString() ?? string.Empty;
        var decoded = System.Net.WebUtility.HtmlDecode(raw);
        sheet.Cell(rowIndex, colIndex).SetValue(XLCellValue.FromObject(decoded));
    }
}

ExcelOverWriter.RegisterOverWriteFunction(new HtmlEntityFunction());
```

The contract:

| Member | Notes |
| --- | --- |
| `string Name { get; }` | The text that follows `#`. The engine matches `#{Name}(` exactly, so keep it short and unambiguous. |
| `Task InvokeAsync(IXLWorksheet sheet, int rowIndex, int colIndex, object?[] args)` | Called once per matching cell. `args` already has `$` symbols resolved (or `null` for unresolved ones). The function should clear or rewrite the cell as appropriate. |

`RegisterOverWriteFunction` appends to a process-wide list, so register at startup before calling `OverWrite` and avoid registering the same function twice.

## See also

* [template-overwrite.md](template-overwrite.md) — how `$symbols` and the symbol converter feed the function arguments.
* [api-reference.md](api-reference.md) — full signatures for `IOverWriteFunction`, `IExcelSymbolConverter`, `ExcelOverWriteCell`.
