using ClosedXML.Excel;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace Excel.Report.PDF
{
    static class PaperSizeMap
    {
        internal static PdfPage AddPage(this PdfDocument pdf, IXLPageSetup ps)
        {
            var page = pdf.AddPage();

            var (wMm, hMm) = GetPaperSize(ps.PaperSize);
            page.Width = wMm;
            page.Height = hMm;

            if (ps.PageOrientation == XLPageOrientation.Landscape) page.Orientation = PdfSharp.PageOrientation.Landscape;

            return page;
        }

        static (XUnit w, XUnit h) GetPaperSize(XLPaperSize s) => s switch
        {
            // US（inch）
            XLPaperSize.LetterPaper or XLPaperSize.LetterSmallPaper or XLPaperSize.NotePaper
                => (XUnit.FromInch(8.5), XUnit.FromInch(11)),
            XLPaperSize.TabloidPaper or XLPaperSize.StandardPaper1
                => (XUnit.FromInch(11), XUnit.FromInch(17)),
            XLPaperSize.LedgerPaper
                => (XUnit.FromInch(17), XUnit.FromInch(11)),
            XLPaperSize.LegalPaper
                => (XUnit.FromInch(8.5), XUnit.FromInch(14)),
            XLPaperSize.StatementPaper
                => (XUnit.FromInch(5.5), XUnit.FromInch(8.5)),
            XLPaperSize.ExecutivePaper
                => (XUnit.FromInch(7.25), XUnit.FromInch(10.5)),
            XLPaperSize.FolioPaper
                => (XUnit.FromInch(8.5), XUnit.FromInch(13)),
            XLPaperSize.StandardPaper
                => (XUnit.FromInch(10), XUnit.FromInch(14)),
            XLPaperSize.StandardPaper2
                => (XUnit.FromInch(9), XUnit.FromInch(11)),
            XLPaperSize.StandardPaper3
                => (XUnit.FromInch(10), XUnit.FromInch(11)),
            XLPaperSize.StandardPaper4
                => (XUnit.FromInch(15), XUnit.FromInch(11)),
            XLPaperSize.UsStandardFanfold
                => (XUnit.FromInch(14.875), XUnit.FromInch(11)),
            XLPaperSize.GermanStandardFanfold
                => (XUnit.FromInch(8.5), XUnit.FromInch(12)),
            XLPaperSize.GermanLegalFanfold
                => (XUnit.FromInch(8.5), XUnit.FromInch(13)),

            // ISO/JIS（mm）
            XLPaperSize.A2Paper
                => (XUnit.FromMillimeter(420), XUnit.FromMillimeter(594)),
            XLPaperSize.A3Paper or XLPaperSize.A3TransversePaper
                => (XUnit.FromMillimeter(297), XUnit.FromMillimeter(420)),
            XLPaperSize.A3ExtraPaper or XLPaperSize.A3ExtraTransversePaper
                => (XUnit.FromMillimeter(322), XUnit.FromMillimeter(445)),
            XLPaperSize.A4Paper or XLPaperSize.A4SmallPaper or XLPaperSize.A4TransversePaper
                => (XUnit.FromMillimeter(210), XUnit.FromMillimeter(297)),
            XLPaperSize.A4PlusPaper
                => (XUnit.FromMillimeter(210), XUnit.FromMillimeter(330)),
            XLPaperSize.A4ExtraPaper
                => (XUnit.FromMillimeter(236), XUnit.FromMillimeter(322)),
            XLPaperSize.A5Paper or XLPaperSize.A5TransversePaper
                => (XUnit.FromMillimeter(148), XUnit.FromMillimeter(210)),
            XLPaperSize.A5ExtraPaper
                => (XUnit.FromMillimeter(174), XUnit.FromMillimeter(235)),
            XLPaperSize.B4Paper or XLPaperSize.IsoB4
                => (XUnit.FromMillimeter(250), XUnit.FromMillimeter(353)),
            XLPaperSize.B5Paper
                => (XUnit.FromMillimeter(176), XUnit.FromMillimeter(250)),
            XLPaperSize.IsoB5ExtraPaper
                => (XUnit.FromMillimeter(201), XUnit.FromMillimeter(276)),
            XLPaperSize.JisB5TransversePaper
                => (XUnit.FromMillimeter(182), XUnit.FromMillimeter(257)),

            // Envelopes and postcards（mm/inch）
            XLPaperSize.DlEnvelope
                => (XUnit.FromMillimeter(110), XUnit.FromMillimeter(220)),
            XLPaperSize.C5Envelope
                => (XUnit.FromMillimeter(162), XUnit.FromMillimeter(229)),
            XLPaperSize.C3Envelope
                => (XUnit.FromMillimeter(324), XUnit.FromMillimeter(458)),
            XLPaperSize.C4Envelope
                => (XUnit.FromMillimeter(229), XUnit.FromMillimeter(324)),
            XLPaperSize.C6Envelope
                => (XUnit.FromMillimeter(114), XUnit.FromMillimeter(162)),
            XLPaperSize.C65Envelope
                => (XUnit.FromMillimeter(114), XUnit.FromMillimeter(229)),
            XLPaperSize.B4Envelope
                => (XUnit.FromMillimeter(250), XUnit.FromMillimeter(353)),
            XLPaperSize.B5Envelope
                => (XUnit.FromMillimeter(176), XUnit.FromMillimeter(250)),
            XLPaperSize.B6Envelope
                => (XUnit.FromMillimeter(176), XUnit.FromMillimeter(125)),
            XLPaperSize.ItalyEnvelope
                => (XUnit.FromMillimeter(110), XUnit.FromMillimeter(230)),
            XLPaperSize.JapaneseDoublePostcard
                => (XUnit.FromMillimeter(200), XUnit.FromMillimeter(148)),
            XLPaperSize.InviteEnvelope
                => (XUnit.FromMillimeter(220), XUnit.FromMillimeter(220)),
            XLPaperSize.MonarchEnvelope
                => (XUnit.FromInch(3.875), XUnit.FromInch(7.5)),
            XLPaperSize.No634Envelope
                => (XUnit.FromInch(3.625), XUnit.FromInch(6.5)),
            XLPaperSize.No9Envelope
                => (XUnit.FromInch(3.875), XUnit.FromInch(8.875)),
            XLPaperSize.No10Envelope
                => (XUnit.FromInch(4.125), XUnit.FromInch(9.5)),
            XLPaperSize.No11Envelope
                => (XUnit.FromInch(4.5), XUnit.FromInch(10.375)),
            XLPaperSize.No12Envelope
                => (XUnit.FromInch(4.75), XUnit.FromInch(11)),
            XLPaperSize.No14Envelope
                => (XUnit.FromInch(5), XUnit.FromInch(11.5)),

            // Large size（inch）
            XLPaperSize.CPaper
                => (XUnit.FromInch(17), XUnit.FromInch(22)),
            XLPaperSize.DPaper
                => (XUnit.FromInch(22), XUnit.FromInch(34)),
            XLPaperSize.EPaper
                => (XUnit.FromInch(34), XUnit.FromInch(44)),

            // Extra/Plus/Transverse（inch or mm）
            XLPaperSize.LetterExtraPaper
                => (XUnit.FromInch(9.275), XUnit.FromInch(12)),
            XLPaperSize.LegalExtraPaper
                => (XUnit.FromInch(9.275), XUnit.FromInch(15)),
            XLPaperSize.TabloidExtraPaper
                => (XUnit.FromInch(11.69), XUnit.FromInch(18)),
            XLPaperSize.LetterTransversePaper
                => (XUnit.FromInch(8.275), XUnit.FromInch(11)),
            XLPaperSize.LetterExtraTransversePaper
                => (XUnit.FromInch(9.275), XUnit.FromInch(12)),
            XLPaperSize.SuperaSuperaA4Paper
                => (XUnit.FromMillimeter(227), XUnit.FromMillimeter(356)),
            XLPaperSize.SuperbSuperbA3Paper
                => (XUnit.FromMillimeter(305), XUnit.FromMillimeter(487)),
            XLPaperSize.LetterPlusPaper
                => (XUnit.FromInch(8.5), XUnit.FromInch(12.69)),

            // Quarto（mm）
            XLPaperSize.QuartoPaper
                => (XUnit.FromMillimeter(215), XUnit.FromMillimeter(275)),

            // default（A4）
            _ => (XUnit.FromMillimeter(210), XUnit.FromMillimeter(297)),
        };
    }
}
