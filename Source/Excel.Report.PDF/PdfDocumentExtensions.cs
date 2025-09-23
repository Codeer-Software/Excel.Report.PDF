using ClosedXML.Excel;
using PdfSharp.Pdf;

namespace Excel.Report.PDF
{
    public static class PdfDocumentExtensions
    {
        public static void AddPages(this PdfDocument pdf, XLWorkbook Workbook)
        {
            using var mem = new MemoryStream();
            Workbook.SaveAs(mem);
            using var openClosedXML = new OpenClosedXML(mem);
            var converter = new PdfRenderer(openClosedXML);
            converter.RenderTo(pdf);
        }
    }
}
