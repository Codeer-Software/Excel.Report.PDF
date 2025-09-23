using PdfSharp.Pdf;

namespace Excel.Report.PDF
{
    public static class ExcelConverter 
    {
        public static int MaxRow = 2000;
        public static int MaxColumn = 256;

        public static MemoryStream ConvertToPdf(string filePath)
        {
            using var openClosedXML = new OpenClosedXML(filePath);
            var converter = new PdfRenderer(openClosedXML);
            using var pdf = new PdfDocument();
            converter.RenderTo(pdf);
            return ToMeoryStream(pdf);
        }

        public static MemoryStream ConvertToPdf(Stream stream)
        {
            using var openClosedXML = new OpenClosedXML(stream);
            var converter = new PdfRenderer(openClosedXML);
            using var pdf = new PdfDocument();
            converter.RenderTo(pdf);
            return ToMeoryStream(pdf);
        }

        public static MemoryStream ConvertToPdf(string filePath, int sheetPosition, PageBreakInfo? pageBreakInfo = null)
        {
            using var openClosedXML = new OpenClosedXML(filePath);
            var converter = new PdfRenderer(openClosedXML);
            using var pdf = new PdfDocument();
            converter.RenderTo(pdf, sheetPosition, pageBreakInfo);
            return ToMeoryStream(pdf);
        }

        public static MemoryStream ConvertToPdf(Stream stream, int sheetPosition, PageBreakInfo? pageBreakInfo = null)
        {
            using var openClosedXML = new OpenClosedXML(stream);
            var converter = new PdfRenderer(openClosedXML);
            using var pdf = new PdfDocument();
            converter.RenderTo(pdf, sheetPosition, pageBreakInfo);
            return ToMeoryStream(pdf);
        }

        public static MemoryStream ConvertToPdf(string filePath, string sheetName, PageBreakInfo? pageBreakInfo = null)
        {
            using var openClosedXML = new OpenClosedXML(filePath);
            var converter = new PdfRenderer(openClosedXML);
            using var pdf = new PdfDocument();
            converter.RenderTo(pdf, openClosedXML.GetSheetPosition(sheetName), pageBreakInfo);
            return ToMeoryStream(pdf);
        }

        public static MemoryStream ConvertToPdf(Stream stream, string sheetName, PageBreakInfo? pageBreakInfo = null)
        {
            using var openClosedXML = new OpenClosedXML(stream);
            var converter = new PdfRenderer(openClosedXML);
            using var pdf = new PdfDocument();
            converter.RenderTo(pdf, openClosedXML.GetSheetPosition(sheetName), pageBreakInfo);
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
