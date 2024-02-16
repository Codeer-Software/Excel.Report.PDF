namespace Excel.Report.PDF
{
    public class ObjectExcelSymbolConverter : IExcelSymbolConverter
    {
        object _obj;
        public ObjectExcelSymbolConverter(object obj) => _obj = obj;

        public async Task<ExcelOverWriteCell?> GetData(string symbol)
        {
            await Task.CompletedTask;
            var prop = _obj.GetType().GetProperty(symbol);
            return prop == null ? null : new ExcelOverWriteCell { Value = prop.GetValue(_obj) };
        }

        public async Task<ExcelOverWriteCell?> GetData(object? element, string elementName, string symbol)
        {
            if (!symbol.StartsWith(elementName + ".")) return await GetData(symbol);
            if (element == null) return new ExcelOverWriteCell();
            var prop = element.GetType().GetProperty(symbol.Substring((elementName + ".").Length));
            return prop == null ? null : new ExcelOverWriteCell { Value = prop.GetValue(element) };
        }
    }
}
