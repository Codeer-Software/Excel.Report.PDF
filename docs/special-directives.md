# Special rendering directives

These directives influence how cells are **drawn** rather than how they are bound to data. Most apply only when the workbook is being converted to PDF (or printed via `Excel.Report.PrintDocument`); they are inert when the workbook is opened in Excel.

A single cell can carry multiple directives by separating them with `|`:

```text
#Empty | #FitColumn   ← also valid
```

Whitespace around the pipe is ignored.

## Page-number directives

These render the current PDF page number into the cell. They are evaluated by the renderer, so the workbook itself never carries the resolved value. Place them in **any column except column A** — column A is reserved for loop directives.

| Directive | Result |
| --- | --- |
| `#Page` | Current page number (1-based). Resolved when the cell is drawn. |
| `#PageCount` | Total number of pages. Resolved as a post-process pass after every page is laid out. |
| `#PageOf("/")` | Renders `current<separator>total`. The separator is the literal between the parentheses (e.g. `"/"`, `" of "`, `"-"`). |

Example:

```text
| Page #Page of #PageCount |
| #PageOf(" / ")           |
```

In a 5-page document, the first page renders as `Page 1 of 5` (or `1 / 5` in the second cell), the second page as `Page 2 of 5`, and so on.

The library inserts a tiny post-processing queue so `#PageCount` and `#PageOf` can be back-filled once total pages are known. There is nothing for you to wire up — it is automatic.

## `#Empty`

```text
| #Empty |
```

Normally only cells containing values (or with a non-default fill, border, etc.) participate in the rendering range calculation. `#Empty` keeps a cell **in the rendering range** without drawing any text:

* The cell still contributes to layout (column widths, row heights, bounds for `#FitColumn`).
* No glyphs are rendered — useful for placeholder cells that you want to influence layout but keep visually empty.

Use this when you want a cell's *presence* (and therefore its borders/fills) to count, but the actual text to remain invisible.

## `#FitColumn`

```text
| #FitColumn |  ← cell A1 only
```

When `A1` contains `#FitColumn`, the renderer scales the entire sheet so the **used column width fills the printable page width** (page width minus left and right margins). This is equivalent to enabling Excel's "Fit to 1 page wide" without the per-page height constraint.

Notes:

* `#FitColumn` only takes effect in cell `A1`.
* It overrides the workbook's manual `Scale` percentage for the affected sheet.
* The library also detects the equivalent Excel setting (`PageSetup.PagesWide > 0`) and applies the same scaling, so you can either set the directive or use Excel's UI.

## Vertical / rotated text

The renderer respects `IXLAlignment.TextRotation`:

* `0` — horizontal (default).
* `1`–`90` — text rotates **counter-clockwise** by the specified angle.
* `91`–`180` — text rotates **clockwise** (Excel's "negative rotation" range).
* `255` — Excel's "Vertical Text" stack: characters are placed top-to-bottom, columns advance left-to-right.

These behaviours are reproduced automatically — no directive is required.

## Number-format-driven hiding

Cells whose number format is set to `";;;"` are intentionally hidden in Excel. The renderer matches that behaviour and skips drawing them, which is useful for staging values you want to use in formulas but never display.

## See also

* [getting-started.md](getting-started.md) — page setup, paper size, and margin handling.
* [multi-page.md](multi-page.md) — combining the page-number directives with `#PagedLoopRows`.
