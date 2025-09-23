namespace Excel.Report.PDF
{
    public static class ExcelConverter 
    {
        public static int MaxRow = 2000;
        public static int MaxColumn = 256;

        public static MemoryStream ConvertToPdf(string filePath)
        {
            using (var converter = new ExcelConverterCore(filePath))
                return converter.ConvertToPdf();
        }

        public static MemoryStream ConvertToPdf(Stream stream)
        {
            using (var converter = new ExcelConverterCore(stream))
                return converter.ConvertToPdf();
        }

        public static MemoryStream ConvertToPdf(string filePath, int sheetPosition, PageBreakInfo? pageBreakInfo = null)
        {
            using (var converter = new ExcelConverterCore(filePath))
                return converter.ConvertToPdf(sheetPosition, pageBreakInfo);
        }

        public static MemoryStream ConvertToPdf(Stream stream, int sheetPosition, PageBreakInfo? pageBreakInfo = null)
        {
            using (var converter = new ExcelConverterCore(stream))
                return converter.ConvertToPdf(sheetPosition, pageBreakInfo);
        }

        public static MemoryStream ConvertToPdf(string filePath, string sheetName, PageBreakInfo? pageBreakInfo = null)
        {
            using (var converter = new ExcelConverterCore(filePath))
                return converter.ConvertToPdf(sheetName, pageBreakInfo);
        }

        public static MemoryStream ConvertToPdf(Stream stream, string sheetName, PageBreakInfo? pageBreakInfo = null)
        {
            using (var converter = new ExcelConverterCore(stream))
                return converter.ConvertToPdf(sheetName, pageBreakInfo);
        }
    }
}
