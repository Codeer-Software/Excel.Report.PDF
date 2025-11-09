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

        static MemoryStream ToPdfMemory(PdfVirtualDocument doc)
        {
            using var pdf = new PdfDocument();

            for (int i = 0; i < doc.PageCount; i++)
            {
                var p = doc.Pages[i];
                var page = pdf.AddPage(p.PageSetup);
                using var gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(page);
                doc.DrawTo(gfx, i);
            }

            return ToMeoryStream(pdf);
        }

        public static MemoryStream ConvertToPdf(Stream stream)
        {
            using var openClosedXML = new OpenClosedXML(stream);
            var converter = new CommonDocumentRender(openClosedXML);
            var document = new PdfVirtualDocument();
            converter.RenderTo(document);
            converter.PostProcessCommands.ExecuteAll();
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
            var converter = new CommonDocumentRender(openClosedXML);
            using var pdf = new PdfDocument();
            var document = new PdfVirtualDocument();
            converter.RenderTo(document, sheetPosition, pageBreakInfo);
            converter.PostProcessCommands.ExecuteAll();
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
            var converter = new CommonDocumentRender(openClosedXML);
            using var pdf = new PdfDocument();
            var document = new PdfVirtualDocument();
            converter.RenderTo(document, openClosedXML.GetSheetPosition(sheetName), pageBreakInfo);
            converter.PostProcessCommands.ExecuteAll();
            return ToPdfMemory(document);
        }

        static MemoryStream ToMeoryStream(PdfDocument pdf)
        {
            var outStream = new MemoryStream();
            pdf.Save(outStream);
            return outStream;
        }
    }
}
