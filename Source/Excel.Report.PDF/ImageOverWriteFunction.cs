using ClosedXML.Excel;

namespace Excel.Report.PDF
{
    class ImageOverWriteFunction : IOverWriteFunction
    {
        public string Name => "Image";

        public async Task InvokeAsync(IXLWorksheet sheet, int rowIndex, int colIndex, object?[] args)
        {
            await Task.CompletedTask;

            Stream? stream = null;
            IDisposable? disposeTarget = null;
            if (0 < args.Length)
            {
                stream = args[0] as Stream;
                if (stream == null)
                {
                    var bin = args[0] as byte[];
                    if (bin != null)
                    {
                        stream = new MemoryStream(bin);
                        disposeTarget = stream;
                    }
                }
            }
            double? widthScale = null;
            if (1 < args.Length)
            {
                if (double.TryParse(args[1]?.ToString() ?? string.Empty, out var v)) widthScale = v;
            }
            double? heightScale = null;
            if (2 < args.Length)
            {
                if (double.TryParse(args[2]?.ToString() ?? string.Empty, out var v)) heightScale = v;
            }

            if (stream == null) return;

            var image = sheet.AddPicture(stream);
            image.MoveTo(sheet.Cell(rowIndex, colIndex), 0, 0);
            if (widthScale != null)
            {
                image.ScaleWidth(widthScale.Value);
            }
            if (heightScale != null)
            {
                image.ScaleHeight(heightScale.Value);
            }
            sheet.Cell(rowIndex, colIndex).SetValue(XLCellValue.FromObject(null));
            disposeTarget?.Dispose();
        }
    }
}
