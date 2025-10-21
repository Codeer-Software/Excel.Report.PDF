using PdfSharp.Pdf;

namespace Excel.Report.PDF
{
    public static class ExcelConverter 
    {
        public static MemoryStream ConvertToPdf(string filePath)
        {
            using var mem = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            return ConvertToPdf(mem);
        }

        public static MemoryStream ConvertToPdf(Stream stream)
        {
            using var openClosedXML = new OpenClosedXML(stream);
            var converter = new PdfRenderer(openClosedXML);
            using var pdf = new PdfDocument();
            converter.RenderTo(pdf);
            converter.PostProcessCommands.ExecuteAll();
            return ToMeoryStream(pdf);
        }

        public static MemoryStream ConvertToPdf(string filePath, int sheetPosition, PageBreakInfo? pageBreakInfo = null)
        {
            using var mem = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            return ConvertToPdf(mem, sheetPosition, pageBreakInfo);
        }

        public static MemoryStream ConvertToPdf(Stream stream, int sheetPosition, PageBreakInfo? pageBreakInfo = null)
        {
            using var openClosedXML = new OpenClosedXML(stream);
            var converter = new PdfRenderer(openClosedXML);
            using var pdf = new PdfDocument();
            converter.RenderTo(pdf, sheetPosition, pageBreakInfo);
            converter.PostProcessCommands.ExecuteAll();
            return ToMeoryStream(pdf);
        }

        public static MemoryStream ConvertToPdf(string filePath, string sheetName, PageBreakInfo? pageBreakInfo = null)
        {
            using var mem = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            return ConvertToPdf(mem, sheetName, pageBreakInfo);
        }

        public static MemoryStream ConvertToPdf(Stream stream, string sheetName, PageBreakInfo? pageBreakInfo = null)
        {
            using var openClosedXML = new OpenClosedXML(stream);
            var converter = new PdfRenderer(openClosedXML);
            using var pdf = new PdfDocument();
            converter.RenderTo(pdf, openClosedXML.GetSheetPosition(sheetName), pageBreakInfo);
            converter.PostProcessCommands.ExecuteAll();
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
