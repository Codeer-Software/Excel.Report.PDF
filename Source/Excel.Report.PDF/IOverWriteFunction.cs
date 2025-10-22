using ClosedXML.Excel;

namespace Excel.Report.PDF
{
    public interface IOverWriteFunction
    {
        public string Name { get; }
        Task InvokeAsync(IXLWorksheet sheet, int rowIndex, int colIndex, object?[] args);
    }
}
