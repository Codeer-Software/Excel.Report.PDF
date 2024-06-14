namespace Excel.Report.PDF
{
    internal class PictureInfo
    {
        internal MemoryStream? Picture { get; set; }
        internal int Index { get; set; }
        internal double X { get; set; }
        internal double Y { get; set; }
        internal double Width { get; set; }
        internal double Height { get; set; }
    }
}
