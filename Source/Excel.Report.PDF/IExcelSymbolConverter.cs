namespace Excel.Report.PDF
{
    public interface IExcelSymbolConverter
    {
        IExcelSymbolConverter CreateChildExcelSymbolConverter(object? obj, string name);
        Task<ExcelOverWriteCell?> GetData(string symbol);
    }
}

namespace Excel.Report.PDF
{
    // Internal fast path for the built-in converter.
    // It avoids creating an async wrapper for every cell; custom converters still use GetData above.
    internal interface ISynchronousExcelSymbolConverter
    {
        // Lets the writer resolve a value without creating a Task.
        bool TryGetData(string symbol, out object? value);
    }
}
