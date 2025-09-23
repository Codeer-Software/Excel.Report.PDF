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
            var converter = new PdfPageRenderer(mem);
            converter.RenderTo(pdf);
        }
    }
}
