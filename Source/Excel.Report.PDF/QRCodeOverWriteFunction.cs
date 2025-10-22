using ClosedXML.Excel;
using QRCoder;

namespace Excel.Report.PDF
{
    class QRCodeOverWriteFunction : IOverWriteFunction
    {
        public string Name => "QR";

        public async Task InvokeAsync(IXLWorksheet sheet, int rowIndex, int colIndex, object?[] args)
        {
            await Task.CompletedTask;

            var qrText = string.Empty;
            if (0 < args.Length)
            {
                qrText = args[0]?.ToString() ?? string.Empty;
            }
            int size = 10;
            if (1 < args.Length)
            {
                if (int.TryParse(args[1]?.ToString() ?? string.Empty, out var v)) size = v;
            }
            if (qrText.StartsWith("\"") && qrText.EndsWith("\""))
            {
                qrText = qrText[1..^1];
            }
            if (!string.IsNullOrEmpty(qrText))
            {
                using var gen = new QRCodeGenerator();
                using var data = gen.CreateQrCode(qrText, QRCodeGenerator.ECCLevel.M);

                var png = new PngByteQRCode(data);
                var bytes = png.GetGraphic(
                    pixelsPerModule: size,
                    darkColorRgba: new byte[] { 0, 0, 0, 255 },
                    lightColorRgba: new byte[] { 255, 255, 255, 255 },
                    drawQuietZones: true);
                using var stream = new MemoryStream(bytes);
                var image = sheet.AddPicture(stream);
                image.MoveTo(sheet.Cell(rowIndex, colIndex), 0, 0);
                sheet.Cell(rowIndex, colIndex).SetValue(XLCellValue.FromObject(null));
            }
        }
    }
}
