# PrintDocument integration

`Excel.Report.PrintDocument` lets you reuse the same Excel rendering pipeline used by `ExcelConverter.ConvertToPdf`, but draw onto a `System.Drawing.Printing.PrintDocument` instead of a PDF page. This is convenient when you need:

* `PrintPreviewDialog` integration on Windows desktop apps (WinForms / WPF host),
* direct printing to an installed printer driver,
* PageSetup-aware output (`PageSetupDialog` produces a `PageSettings` you can convert into the renderer's coordinate system).

> The package is **Windows-only**. It depends on `System.Drawing.Common` and is annotated with `[SupportedOSPlatform("windows")]`.

## Install

```powershell
PM> Install-Package Excel.Report.PrintDocument
```

The package transitively pulls in `Excel.Report.PDF`, so you get `ExcelConverter`, `ExcelOverWriter`, and the symbol/function APIs at the same time.

## Bind a workbook to a `PrintDocument`

```csharp
using Excel.Report.PrintDocument;
using System.Drawing.Printing;

var doc = new PrintDocument();

// Hook the workbook into the document. Optionally override the page setup with one
// derived from the user's PageSetupDialog selection.
ExcelPrintDocumentBinder.Bind(doc, "report.xlsx");

// Use any standard PrintDocument flow:
using var preview = new PrintPreviewDialog { Document = doc, Width = 1000, Height = 800 };
preview.ShowDialog();

// or doc.Print();
```

`Bind` wires up two event handlers on `PrintDocument`:

* `PrintPage` — draws each rendered page onto the printer graphics surface and sets `e.HasMorePages` while pages remain.
* `EndPrint` — detaches both handlers so you can call `Bind` again later without leaking subscriptions.

Each call to `Bind` is independent — pass the next workbook to the same `PrintDocument` instance and it works again.

### Stream input

`Bind` accepts a `Stream` overload, which is convenient when the workbook has just been produced by the template overwrite pipeline:

```csharp
using var book = new XLWorkbook("template.xlsx");
await book.OverWrite(new ObjectExcelSymbolConverter(data));

using var ms = new MemoryStream();
book.SaveAs(ms);
ExcelPrintDocumentBinder.Bind(doc, ms);
```

## Custom page setup

By default the renderer reads `PageSetup` from the first worksheet of the workbook. To override it (for example after the user picks paper size from `PageSetupDialog`), pass a `PrintPageSetup`:

```csharp
var setup = new PrintPageSetup
{
    WidthPoint  = PrintPageSetup.MmToPoint(210), // A4 width
    HeightPoint = PrintPageSetup.MmToPoint(297), // A4 height
    Margins = new PrintMargins
    {
        LeftPoint   = PrintPageSetup.MmToPoint(15),
        RightPoint  = PrintPageSetup.MmToPoint(15),
        TopPoint    = PrintPageSetup.MmToPoint(20),
        BottomPoint = PrintPageSetup.MmToPoint(20),
    }
};

ExcelPrintDocumentBinder.Bind(doc, "report.xlsx", setup);
```

You can also retrieve the current page setup from a workbook for further tweaking:

```csharp
var current = ExcelPrintDocumentBinder.GetPrintPageSetup("report.xlsx");
current.Margins.LeftPoint = PrintPageSetup.MmToPoint(10);
ExcelPrintDocumentBinder.Bind(doc, "report.xlsx", current);
```

`PrintPageSetup.MmToPoint` is provided as a small helper because page metrics on `PrintDocument` use 1/72 inch (point) units, while end users typically think in millimetres or inches.

## Worked sample

A complete WinForms host is included in the repository:

* [`Source/TestWinFormsApp/MainForm.cs`](../Source/TestWinFormsApp/MainForm.cs) — file picker + preview + page setup wiring
* [`Source/TestWinFormsApp/Program.cs`](../Source/TestWinFormsApp/Program.cs) — installs `WindowsInstalledFontResolver` so PdfSharp can render any installed Windows font

Run the project under Windows to preview any `.xlsx` file end-to-end through the same pipeline used for PDF output.

## Font resolver

The `IFontResolver` requirement still applies — `PrintDocument` uses GDI+ fonts directly, but the underlying renderer routes through the same shared infrastructure as PDF output.

If you only target Windows desktop, [`WindowsInstalledFontResolver`](../Source/TestWinFormsApp/WindowsInstalledFontResolver.cs) is a good drop-in implementation that reads from the system font registry.
