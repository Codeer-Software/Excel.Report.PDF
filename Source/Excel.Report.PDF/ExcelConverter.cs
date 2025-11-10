using Excel.Report.PrintDocument;
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
            var converter = new VirtualRender(openClosedXML);
            var document = new PdfVirtualDocument();
            converter.RenderTo(document);
            return ToPdfMemory(document);
        }

        public static MemoryStream ConvertToPdf(string filePath, int sheetPosition, PageBreakInfo? pageBreakInfo = null)
        {
            using var mem = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            return ConvertToPdf(mem, sheetPosition, pageBreakInfo);
        }

        public static MemoryStream ConvertToPdf(Stream stream, int sheetPosition, PageBreakInfo? pageBreakInfo = null)
        {
            using var openClosedXML = new OpenClosedXML(stream);
            var converter = new VirtualRender(openClosedXML);
            var document = new PdfVirtualDocument();
            converter.RenderTo(document, sheetPosition, pageBreakInfo);
            return ToPdfMemory(document);
        }

        public static MemoryStream ConvertToPdf(string filePath, string sheetName, PageBreakInfo? pageBreakInfo = null)
        {
            using var mem = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            return ConvertToPdf(mem, sheetName, pageBreakInfo);
        }

        public static MemoryStream ConvertToPdf(Stream stream, string sheetName, PageBreakInfo? pageBreakInfo = null)
        {
            using var openClosedXML = new OpenClosedXML(stream);
            var converter = new VirtualRender(openClosedXML);
            var document = new PdfVirtualDocument();
            converter.RenderTo(document, openClosedXML.GetSheetPosition(sheetName), pageBreakInfo);
            return ToPdfMemory(document);
        }

        static MemoryStream ToPdfMemory(PdfVirtualDocument doc)
        {
            using var pdf = new PdfDocument();

            foreach(var virtualPage in doc.Pages)
            {
                var page = pdf.AddPage(virtualPage.PageSetup);
                using var gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(page);
                virtualPage.DrawTo(gfx);
            }

            var outStream = new MemoryStream();
            pdf.Save(outStream);
            return outStream;
        }
    }
}
