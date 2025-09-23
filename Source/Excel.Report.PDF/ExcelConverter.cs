using PdfSharp.Pdf;

namespace Excel.Report.PDF
{
    public static class ExcelConverter 
    {
        public static int MaxRow = 2000;
        public static int MaxColumn = 256;

        public static MemoryStream ConvertToPdf(string filePath)
        {
            using var converter = new ExcelConverterCore(filePath);
            using var pdf = new PdfDocument();
            converter.ConvertToPdf(pdf);
            return ToMeoryStream(pdf);
        }

        public static MemoryStream ConvertToPdf(Stream stream)
        {
            using var converter = new ExcelConverterCore(stream);
            using var pdf = new PdfDocument();
            converter.ConvertToPdf(pdf);
            return ToMeoryStream(pdf);
        }

        public static MemoryStream ConvertToPdf(string filePath, int sheetPosition, PageBreakInfo? pageBreakInfo = null)
        {
            using var converter = new ExcelConverterCore(filePath);
            using var pdf = new PdfDocument();
            converter.ConvertToPdf(pdf, sheetPosition, pageBreakInfo);
            return ToMeoryStream(pdf);
        }

        public static MemoryStream ConvertToPdf(Stream stream, int sheetPosition, PageBreakInfo? pageBreakInfo = null)
        {
            using var converter = new ExcelConverterCore(stream);
            using var pdf = new PdfDocument();
            converter.ConvertToPdf(pdf, sheetPosition, pageBreakInfo);
            return ToMeoryStream(pdf);
        }

        public static MemoryStream ConvertToPdf(string filePath, string sheetName, PageBreakInfo? pageBreakInfo = null)
        {
            using var converter = new ExcelConverterCore(filePath);
            using var pdf = new PdfDocument();
            converter.ConvertToPdf(pdf, converter.OpenClosedXML.GetSheetPosition(sheetName), pageBreakInfo);
            return ToMeoryStream(pdf);
        }

        public static MemoryStream ConvertToPdf(Stream stream, string sheetName, PageBreakInfo? pageBreakInfo = null)
        {
            using var converter = new ExcelConverterCore(stream);
            using var pdf = new PdfDocument();
            converter.ConvertToPdf(pdf, converter.OpenClosedXML.GetSheetPosition(sheetName), pageBreakInfo);
            return ToMeoryStream(pdf);
        }

        static MemoryStream ToMeoryStream(PdfDocument pdf)
        {
            var outStream = new MemoryStream();
            pdf.Save(outStream);
            return outStream;
        }
    }
}
