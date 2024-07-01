using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace Excel.Report.PDF
{

    static class PdfDocumentExtensions
    {
        internal static PdfPage AddPageEx(this PdfDocument pdf)
        {
            var page = pdf.AddPage();
            page.Width = XUnit.FromMillimeter(210);
            page.Height = XUnit.FromMillimeter(297);
            return page;
        }
    }
}
