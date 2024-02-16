# Excel.Report.PDF

## Features ...
Convert Excel to PDF.<br>
<img src="Image/SampleExcelToPDF.png" width="800">

Overwrite Excel with data according to the symbol.(And convert to Excel to PDF)<br>
<img src="Image/SampleQuotation.png" width="800">

## Getting Started
Excel.Report.PDF from NuGet.

    PM> Install-Package Excel.Report.PDF

First you need to implement IFontResolver.
This is a PDFsharp feature used internally.
I think you can understand if you refer to the test project.
->Source/Test/Test
```csharp
public class CustomFontResolver : IFontResolver
{
    public byte[] GetFont(string faceName)
        //Implement so that you can get as many fonts as you need.
        => faceName.EndsWith("#b") ? Resources.NotoSansJP_ExtraBold : Resources.NotoSansJP_Regular;

    public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        var faceName = familyName; 
        if (isBold) faceName += "#b";
        return new FontResolverInfo(faceName);
    }
}
```
```csharp
GlobalFontSettings.FontResolver = new CustomFontResolver();
```

Next, you can convert by specifying Excel. The first argument is the Excel path or Stream, and the second argument is the number or name of the target sheet.
```csharp
using var outStream = ExcelConverter.ConvertToPdf(workbookPath, 1);
File.WriteAllBytes(pdfPath, outStream.ToArray());
```
## Overwrite Excel
There is also a function to overwrite Excel.
First, create an Excel file according to the rules.

### 1. [$]
Write the string you want to convey to the program after the $ sign.<br>
<img src="Image/SymbolDollar.png">

### 2. [#LoopRow($elements, elementName, copyRowCount)]
It can be specified in column A. Copies the specified row as many times as copyRowCount and adds the row.
In the copied column, you can specify the symbol of the repeating element using elementName.

| Name |  |
| ---- | ---- |
| $elements | A loop element. Specifies the symbol returned by IEnumerable. |
| elementName | The name of a repeating element used within a row. |
| copyRowCount | Number of rows to copy. |

<br>
<img src="Image/SymbolRowCopy.png">
Next, programmatically overwrite this Excel.
Pass IExcelSymbolConverter to the OverWrite method.
This sample uses ObjectExcelSymbolConverter, which is one of the implementations of IExcelSymbolConverter provided by Excel.Report.PDF.
IExcelSymbolConverter is easy to implement, so please try implementing it according to your situation.
In that case, the implementation of ObjectExcelSymbolConverter will be helpful.

```csharp
//Sample data.
var data = new Quotation 
{
    Title = "宴会時の食材",
    Client = "エクセルコンサルティング株式会社",
    PersonInCharge = "大谷正一"
};
data.Details.Add(new()
{
    Title = "鯛",
    Detail = "新鮮",
    Price = 10000,
    Discount = 0,
});
data.Details.Add(new()
{
    Title = "鰤",
    Detail = "新鮮",
    Price = 20000,
    Discount = 0,
});
data.Details.Add(new()
{
    Title = "ハマチ",
    Detail = "ご奉仕品",
    Price = 30000,
    Discount = 2000,
});
data.Details.Add(new()
{
    Title = "蛸",
    Detail = "ご奉仕品",
    Price = 40000,
    Discount = 1000,
});

using var book = new XLWorkbook(filePath);
var symbolConverter = new ObjectExcelSymbolConverter(data);
await book.Worksheet(1).OverWrite(new ObjectExcelSymbolConverter(data));

// Convert Excel to PDF
var newStream = new MemoryStream();
book.SaveAs(newStream);
using var outStream = ExcelConverter.ConvertToPdf(newStream, 1);
File.WriteAllBytes(pdfPath, outStream.ToArray());
```

