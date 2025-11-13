using Excel.Report.PDF;
using System.Drawing.Printing;
using System.Runtime.Versioning;

namespace Excel.Report.PrintDocument
{
    [SupportedOSPlatform("windows")]
    public class ExcelPrintDocumentBinder
    {
        public static void Bind(System.Drawing.Printing.PrintDocument doc, string filePath, PrintPageSetup? setup = null)
        {
            using var mem = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            Bind(doc, mem, setup);
        }

        public static void Bind(System.Drawing.Printing.PrintDocument doc, Stream stream, PrintPageSetup? setup = null)
        {
            var pageSetup = setup?.ToPageSetup();

            using var openClosedXML = new OpenClosedXML(stream);
            var converter = new VirtualRender(openClosedXML);
            var document = new PrintVirtualDocument();
            converter.RenderTo(document, pageSetup);

            int pageIndex = 0;
            void OnPrintPage(object? sender, PrintPageEventArgs e)
            {
                e.Graphics!.PageUnit = System.Drawing.GraphicsUnit.Point;
                document.DrawTo(e.Graphics!, pageIndex);
                pageIndex++;
                e.HasMorePages = pageIndex < document.PageCount;
            }
            void OnEndPrint(object? sender, PrintEventArgs e)
            {
                doc.PrintPage -= OnPrintPage;
                doc.EndPrint -= OnEndPrint;
            }
            doc.PrintPage += OnPrintPage;
            doc.EndPrint += OnEndPrint;
        }

        public static PrintPageSetup GetPrintPageSetup(string filePath)
        {
            using var mem = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            return GetPrintPageSetup(mem);
        }

        public static PrintPageSetup GetPrintPageSetup(Stream stream)
        {
            using var openClosedXML = new OpenClosedXML(stream);
            var sheet = openClosedXML.Workbook.Worksheets.First();
            var pageSetup = sheet.PageSetup;
            return new PrintPageSetup().FromIXLPageSetup(pageSetup);
        }
    }
}