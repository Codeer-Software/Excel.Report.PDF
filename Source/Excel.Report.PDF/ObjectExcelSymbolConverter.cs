namespace Excel.Report.PDF
{
    public class ObjectExcelSymbolConverter : IExcelSymbolConverter, ISynchronousExcelSymbolConverter
    {
        // Caches property getters because the same symbols are often used many times.
        sealed class PropertyAccessor
        {
            public PropertyAccessor(bool exists, Func<object, object?> getter)
            {
                Exists = exists;
                Getter = getter;
            }

            public bool Exists { get; }
            public Func<object, object?> Getter { get; }
        }

        static readonly PropertyAccessor MissingProperty = new(false, _ => null); // Cache missing properties too.
        static readonly System.Collections.Concurrent.ConcurrentDictionary<(Type Type, string Name), PropertyAccessor> PropertyAccessors = new(); // Reuse getters across rows.
        static readonly Task<ExcelOverWriteCell?> MissingData = Task.FromResult<ExcelOverWriteCell?>(null); // Reuse the common missing result.

        readonly object? _obj;
        readonly string _name = string.Empty;
        public ObjectExcelSymbolConverter(object? obj) => _obj = obj;
        
        ObjectExcelSymbolConverter(object? obj, string name)
        {
            _obj = obj;
            _name = name;
        }

        public IExcelSymbolConverter CreateChildExcelSymbolConverter(object? obj, string name)
            => new ObjectExcelSymbolConverter(obj, name);

        public Task<ExcelOverWriteCell?> GetData(string symbol)
        {
            // Keep the public async API while using the faster synchronous lookup internally.
            if (!TryGetData(symbol, out var value)) return MissingData;
            return Task.FromResult<ExcelOverWriteCell?>(new ExcelOverWriteCell { Value = value });
        }

        public bool TryGetData(string symbol, out object? value)
        {
            value = null;

            if (_obj == null) return false;

            var propertyName = symbol;
            if (!string.IsNullOrEmpty(_name))
            {
                var prefix = _name + ".";
                if (!symbol.StartsWith(prefix, StringComparison.Ordinal)) return false;
                propertyName = symbol.Substring(prefix.Length);
            }

            return TryGetPropertyValue(_obj, propertyName, out value);
        }

        public Task<ExcelOverWriteCell?> GetData(object? element, string elementName, string symbol)
        {
            if (_obj == null) return MissingData;

            var prefix = elementName + ".";
            if (!symbol.StartsWith(prefix, StringComparison.Ordinal)) return MissingData;
            if (element == null)
                return Task.FromResult<ExcelOverWriteCell?>(new ExcelOverWriteCell());

            return TryGetPropertyValue(element, symbol.Substring(prefix.Length), out var value)
                ? Task.FromResult<ExcelOverWriteCell?>(new ExcelOverWriteCell { Value = value })
                : MissingData;
        }

        static bool TryGetPropertyValue(object target, string propertyName, out object? value)
        {
            // Create each type/property lookup only once.
            var accessor = PropertyAccessors.GetOrAdd(
                (target.GetType(), propertyName),
                static key => CreatePropertyAccessor(key.Type, key.Name));

            if (!accessor.Exists)
            {
                value = null;
                return false;
            }

            value = accessor.Getter(target);
            return true;
        }

        static PropertyAccessor CreatePropertyAccessor(Type type, string propertyName)
        {
            var property = type.GetProperty(propertyName);
            if (property == null || property.GetMethod == null)
                return MissingProperty;

            try
            {
                // Compiled getters avoid reflection on every output cell.
                var target = System.Linq.Expressions.Expression.Parameter(typeof(object), "target");
                var typedTarget = System.Linq.Expressions.Expression.Convert(target, type);
                var propertyValue = System.Linq.Expressions.Expression.Property(typedTarget, property);
                var boxedValue = System.Linq.Expressions.Expression.Convert(propertyValue, typeof(object));
                var getter = System.Linq.Expressions.Expression.Lambda<Func<object, object?>>(boxedValue, target).Compile();
                return new PropertyAccessor(true, getter);
            }
            catch (ArgumentException)
            {
                // Fall back to reflection for properties that cannot be compiled.
                return new PropertyAccessor(true, property.GetValue);
            }
        }

    }
}
