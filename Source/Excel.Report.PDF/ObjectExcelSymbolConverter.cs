namespace Excel.Report.PDF
{
    public class ObjectExcelSymbolConverter : IExcelSymbolConverter
    {
        object? _obj;
        string _name = string.Empty;
        public ObjectExcelSymbolConverter(object? obj) => _obj = obj;
        
        ObjectExcelSymbolConverter(object? obj, string name)
        {
            _obj = obj;
            _name = name;
        }

        public IExcelSymbolConverter CreateChildExcelSymbolConverter(object? obj, string name)
            => new ObjectExcelSymbolConverter(obj, name);

        public async Task<ExcelOverWriteCell?> GetData(string symbol)
        {
            await Task.CompletedTask;

            if (_obj == null)
                return null;

            if(!string.IsNullOrEmpty(_name))
                return await GetData(_obj, _name, symbol);
            var prop = _obj.GetType().GetProperty(symbol);
            return prop == null ? null : new ExcelOverWriteCell { Value = prop.GetValue(_obj) };
        }

        public async Task<ExcelOverWriteCell?> GetData(object? element, string elementName, string symbol)
        {
            await Task.CompletedTask;

            if (_obj == null)
                return null;
            if (!symbol.StartsWith(elementName + ".")) return null;
            if (element == null) return new ExcelOverWriteCell();
            var prop = element.GetType().GetProperty(symbol.Substring((elementName + ".").Length));
            return prop == null ? null : new ExcelOverWriteCell { Value = prop.GetValue(element) };
        }

    }
}
