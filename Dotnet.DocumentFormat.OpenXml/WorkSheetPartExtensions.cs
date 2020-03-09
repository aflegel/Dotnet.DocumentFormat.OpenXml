using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Dotnet.DocumentFormat.OpenXml
{
	public static partial class WorksheetPartExtensions
	{
		public static Row AddRow(this WorksheetPart workSheet)
		{
			// Get the sheetData cell table.
			var sheetData = workSheet.Worksheet.GetFirstChild<SheetData>();

			// Add a row to the cell table.
			var row = new Row() { RowIndex = new UInt32Value((uint)sheetData.ChildElements.Count + 1) };
			sheetData.Append(row);

			return row;
		}
	}
}
