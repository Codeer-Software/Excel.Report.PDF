using Excel.Report.PDF;
using System.Drawing.Printing;
using System.Runtime.Versioning;

namespace Excel.Report.PrintDocument
{
    [SupportedOSPlatform("windows")]
    public class ExcelPrintDocumentBinder
    {
        //TODO GetPrintSettings

        public static void Bind(System.Drawing.Printing.PrintDocument doc, string filePath)
        {
            using var mem = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            Bind(doc, mem);
        }

        public static void Bind(System.Drawing.Printing.PrintDocument doc, Stream stream)
        {
            using var openClosedXML = new OpenClosedXML(stream);
            var converter = new VirtualRender(openClosedXML);
            var document = new PrintVirtualDocument();
            converter.RenderTo(document);
            converter.PostProcessCommands.ExecuteAll();

            //TODO Orientation and margin settings
            int pageIndex = 0;
            void OnPrintPage(object? sender, PrintPageEventArgs e)
            {
                //TODO add initilize actions.
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
    }
}