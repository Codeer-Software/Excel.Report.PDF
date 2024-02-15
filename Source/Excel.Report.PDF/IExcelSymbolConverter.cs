namespace Excel.Report.PDF
{
    public interface IExcelSymbolConverter
    {
        Task<ExcelOverWriteCell?> GetData(string symbol);
        Task<ExcelOverWriteCell?> GetData(object? element, string elementName, string symbol);
    }
}
