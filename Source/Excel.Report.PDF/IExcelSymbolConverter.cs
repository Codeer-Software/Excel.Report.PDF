namespace Excel.Report.PDF
{
    public interface IExcelSymbolConverter
    {
        IExcelSymbolConverter CreateChildExcelSymbolConverter(object? obj, string name);
        Task<ExcelOverWriteCell?> GetData(string symbol);
    }
}
