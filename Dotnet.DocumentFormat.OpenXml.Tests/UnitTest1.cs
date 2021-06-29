using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;
using Dotnet.DocumentFormat.OpenXml;
using System.Collections.Generic;
using System.Linq;

namespace Dotnet.DocumentFormatOpenXml.Tests
{
	public class UnitTest1
	{
		protected List<string> xlHeaders = new List<string> {
			"ID",
			"Customer Transaction ID",
			"Invoice Number",
			"Amount",
			"Date",
			"Transaction Type"
		};

		[Fact]
		public void Test1()
		{
			//DateTime dateRange;
			//int finalPort;
			var sampleData = new List<object[]>();

			using var fileStream = new MemoryStream();

			var simplyExportDocument = SpreadsheetDocument.Create(fileStream, SpreadsheetDocumentType.Workbook);

			// Add a WorkbookPart to the document.
			var workbookpart = simplyExportDocument.AddWorkBook();

			// Add a WorksheetPart to the WorkbookPart.
			var worksheetPart = workbookpart.AddWorkSheet("Invoices");

			// Get the sheetData cell table.
			var spreadsheetRow = worksheetPart.AddRow();

			xlHeaders.ForEach(header => spreadsheetRow.Append(new Cell() { DataType = CellValues.String, CellValue = new CellValue(header) }));

			sampleData.ForEach(currentRow =>
			{
				var currentSheetRow = worksheetPart.AddRow();

				currentRow.SelectMany(row => xlHeaders, (rowData, header) => rowData).ToList()
				.ForEach(rowData => currentSheetRow.Append(new Cell() { DataType = SpreadsheetDocumentExtensions.GetDatatype(rowData), CellValue = new CellValue(rowData.ToString()) }));

			});

			workbookpart.Workbook.Save();
		}
	}
}
