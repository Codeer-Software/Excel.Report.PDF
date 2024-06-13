using ClosedXML.Excel;

namespace Excel.Report.PDF
{
    class CellInfo
    {
        internal CellInfo? MergedFirstCell { get; set; }
        internal CellInfo? MergedLastCell { get; set; }
        internal string Text { get; set; } = string.Empty;

        internal string BackColor { get; set; } = string.Empty;
        internal string ForeColor { get; set; } = string.Empty;
        internal double X { get; set; }
        internal double Y { get; set; }
        internal double Width { get; set; }
        internal double Height { get; set; }
        internal double MergedWidth { get; set; }
        internal double MergedHeight { get; set; }

        internal IXLCell? Cell { get; set; }

        internal List<PictureInfo> Pictures { get; set; } = new();
    }
}
