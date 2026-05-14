# Getting started

This guide walks through installing `Excel.Report.PDF`, registering a font resolver, and using every conversion overload exposed by `ExcelConverter`.

> Looking for the template / data-binding side of the library? Continue with [template-overwrite.md](template-overwrite.md) after finishing this page.

## 1. Install

```powershell
PM> Install-Package Excel.Report.PDF
```

| Property | Value |
| --- | --- |
| Target framework | `net6.0` |
| Runs on | Windows / Linux / macOS |
| External processes | None — no Office, no COM |

The package transitively depends on:

* [ClosedXML](https://www.nuget.org/packages/ClosedXML) — Excel parsing & overwrite engine
* [PdfSharp](https://www.nuget.org/packages/PdfSharp) — PDF rendering
* [QRCoder](https://www.nuget.org/packages/QRCoder) — backing for the `#QR` directive

## 2. Register a font resolver

PdfSharp does not bundle fonts. You must register an `IFontResolver` **before** the first call to `ExcelConverter.ConvertToPdf`. Otherwise PdfSharp throws when it needs to resolve a font.

### Option A — embed fonts as resources

Recommended for cross-platform deployments where you want predictable rendering.

```csharp
using PdfSharp.Fonts;

public class CustomFontResolver : IFontResolver
{
    public byte[] GetFont(string faceName)
        => faceName.EndsWith("#b") ? Resources.NotoSansJP_ExtraBold
                                   : Resources.NotoSansJP_Regular;

    public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        var faceName = familyName;
        if (isBold) faceName += "#b";
        return new FontResolverInfo(faceName);
    }
}

GlobalFontSettings.FontResolver = new CustomFontResolver();
```

The `#b` suffix is just a convention used by this resolver to encode "bold" into the face name; you can adopt any naming scheme as long as `ResolveTypeface` and `GetFont` agree.

### Option B — load installed Windows fonts

For Windows-only desktop apps, you can resolve fonts from the system registry. A complete reference implementation is available at:

* [`Source/TestWinFormsApp/WindowsInstalledFontResolver.cs`](../Source/TestWinFormsApp/WindowsInstalledFontResolver.cs)

It supports bold/italic style suffixes, falls back through a configurable family list, and caches font bytes per face name.

```csharp
GlobalFontSettings.FontResolver = new WindowsInstalledFontResolver(
    "Yu Gothic UI", "Segoe UI", "Meiryo UI");
```

### Once is enough

`GlobalFontSettings.FontResolver` is a process-wide singleton. Set it once during startup; subsequent assignments throw. In tests, guard the assignment:

```csharp
if (GlobalFontSettings.FontResolver == null)
    GlobalFontSettings.FontResolver = new CustomFontResolver();
```

## 3. Convert Excel to PDF

The `ExcelConverter` static class exposes one entry point with multiple overloads.

| Overload | Behaviour |
| --- | --- |
| `ConvertToPdf(string path)` | Convert **every sheet** in the workbook. |
| `ConvertToPdf(Stream stream)` | Stream-based equivalent of the above. |
| `ConvertToPdf(string path, int sheetPosition)` | Convert a single sheet. `sheetPosition` is **1-based** (matches ClosedXML). |
| `ConvertToPdf(Stream stream, int sheetPosition)` | Stream + sheet position. |
| `ConvertToPdf(string path, string sheetName)` | Convert a single sheet by name. |
| `ConvertToPdf(Stream stream, string sheetName)` | Stream + sheet name. |

All overloads return a `MemoryStream` containing the PDF bytes. The caller owns the returned stream — wrap it with `using`.

```csharp
using var pdf = ExcelConverter.ConvertToPdf("report.xlsx");
File.WriteAllBytes("report.pdf", pdf.ToArray());
```

### Stream input

When you already have the workbook in memory (e.g., after a template overwrite), pass the stream directly:

```csharp
using var book = new XLWorkbook("template.xlsx");
// ... populate ...

using var ms = new MemoryStream();
book.SaveAs(ms);                       // save into memory
using var pdf = ExcelConverter.ConvertToPdf(ms, 1);
File.WriteAllBytes("out.pdf", pdf.ToArray());
```

`ExcelConverter` resets the stream position internally, so you do not need to seek before passing it in.

## 4. Verify the output

Reproduced layout features include:

* Page setup: paper size, orientation, margins, header/footer margins, manual scaling, fit-to-width, centering
* Page breaks (manual `RowBreaks` / `ColumnBreaks` produce additional PDF pages)
* Cell fills, fonts (size / bold / italic / underline / colour incl. theme tints)
* Borders with Excel-style precedence on shared edges, including `Double` borders
* Text alignment, wrapping, multi-line, vertical text (`TextRotation = 255`), rotated text (1°–179°)
* Embedded pictures (positioned by their top-left cell anchor)

If a feature renders unexpectedly, see [special-directives.md](special-directives.md) for hints on `#Empty` and `#FitColumn`.

## 5. Where next?

* Use Excel as a template language → [template-overwrite.md](template-overwrite.md)
* Split long lists across pages → [multi-page.md](multi-page.md)
* Add dynamic images / QR codes → [built-in-functions.md](built-in-functions.md)
* Print or preview instead of writing PDF → [print-document.md](print-document.md)
