using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Dotnet.DocumentFormat.OpenXml
{
	public static class SpreadsheetDocumentExtensions
	{
		public enum StyleIndex
		{
			Default,
			Accounting,
			Date
		}

		public static CellValues GetDatatype(object dataValue) => dataValue.GetType() == typeof(string)
				? CellValues.String
				: dataValue.GetType() == typeof(DateTime)
				? CellValues.Date
				: dataValue.GetType() == typeof(bool) ? CellValues.Boolean : CellValues.Number;

		public static WorkbookPart AddWorkBook(this SpreadsheetDocument document)
		{
			var part = document.AddWorkbookPart();
			part.Workbook = new Workbook();

			// Add Sheets to the Workbook.
			document.WorkbookPart.Workbook.AppendChild(new Sheets());
			document.WorkbookPart.AddStylePart();

			return part;
		}
	}
}
