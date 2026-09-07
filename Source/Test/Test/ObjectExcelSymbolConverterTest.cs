using ClosedXML.Excel;
using Excel.Report.PDF;

namespace Test
{
    // Covers cached object-symbol resolution and the custom async path to prevent performance changes from altering behavior.
    public class ObjectExcelSymbolConverterTest
    {
        class Model
        {
            public string Text { get; set; } = string.Empty;
            public int Number { get; set; }
            public string? NullText { get; set; }
            public string Computed => $"{Text}:{Number}";
        }

        class Owner
        {
            public Model Child { get; set; } = new();
        }

        class AsyncConverter : IExcelSymbolConverter
        {
            public int CallCount { get; private set; }

            public IExcelSymbolConverter CreateChildExcelSymbolConverter(object? obj, string name) => this;

            public Task<ExcelOverWriteCell?> GetData(string symbol)
            {
                CallCount++;
                return Task.FromResult<ExcelOverWriteCell?>(
                    symbol == "Value" ? new ExcelOverWriteCell { Value = "async" } : null);
            }
        }

        [Test]
        public async Task RepeatedPropertyLookupReturnsTheSameValue()
        {
            var converter = new ObjectExcelSymbolConverter(new Model { Text = "hello" });

            for (var i = 0; i < 100; i++)
            {
                var result = await converter.GetData("Text");
                result.IsNotNull();
                result!.Value.Is("hello");
            }
        }

        [Test]
        public async Task MissingPropertyReturnsNull()
        {
            var converter = new ObjectExcelSymbolConverter(new Model());

            var result = await converter.GetData("DoesNotExist");

            result.IsNull();
        }

        [Test]
        public async Task NullPropertyReturnsAWriteCellWithNullValue()
        {
            var converter = new ObjectExcelSymbolConverter(new Model { NullText = null });

            var result = await converter.GetData("NullText");

            result.IsNotNull();
            result!.Value.IsNull();
        }

        [Test]
        public async Task ComputedPropertyIsResolved()
        {
            var converter = new ObjectExcelSymbolConverter(new Model { Text = "hello", Number = 42 });

            var result = await converter.GetData("Computed");

            result.IsNotNull();
            result!.Value.Is("hello:42");
        }

        [Test]
        public async Task ChildConverterResolvesNestedProperty()
        {
            var owner = new Owner { Child = new Model { Text = "nested" } };
            var converter = new ObjectExcelSymbolConverter(owner);
            var child = converter.CreateChildExcelSymbolConverter(owner.Child, "item");

            var result = await child.GetData("item.Text");
            var unrelated = await child.GetData("other.Text");

            result.IsNotNull();
            result!.Value.Is("nested");
            unrelated.IsNull();
        }

        [Test]
        public async Task CustomAsyncConverterUsesTheAsyncPath()
        {
            using var book = new XLWorkbook();
            var sheet = book.AddWorksheet("Sheet1");
            sheet.Cell(1, 1).SetValue("$Value");
            var converter = new AsyncConverter();

            await sheet.OverWrite(converter);

            sheet.Cell(1, 1).GetText().Is("async");
            converter.CallCount.Is(1);
        }
    }
}
