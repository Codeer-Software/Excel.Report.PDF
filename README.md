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
        => Resources.NotoSansJP_VariableFont_wght;

    public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        => new FontResolverInfo(familyName);
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
