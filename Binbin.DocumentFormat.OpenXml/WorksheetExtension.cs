using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Binbin.DocumentFormat.OpenXml
{
    public static class WorksheetExtension
    {
        public static void UpdateCell(this Worksheet worksheet, uint rowIndex, string columnName, CellValues dataType, string text)
        {
            var cell = SpreadsheetDocumentExtension.GetCell(worksheet, rowIndex, columnName);
            cell.CellValue = new CellValue(text);
            cell.DataType = new EnumValue<CellValues>(dataType);
        }
    }
}