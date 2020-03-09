using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Dotnet.DocumentFormat.OpenXml
{
	public static partial class WorkbookPartExtensions
	{
		public static WorksheetPart AddWorkSheet(this WorkbookPart workBook, string sheetName)
		{
			// Add a WorksheetPart to the WorkbookPart.
			var sheetPart = workBook.AddNewPart<WorksheetPart>();
			sheetPart.Worksheet = new Worksheet(new SheetData());

			// Append a new worksheet and associate it with the workbook.
			var sheet = new Sheet() { Id = workBook.GetIdOfPart(sheetPart), SheetId = new UInt32Value((uint)workBook.Parts.Count() + 1), Name = sheetName };
			workBook.Workbook.Sheets.Append(sheet);

			return sheetPart;
		}

		public static void AddStylePart(this WorkbookPart workBook)
		{
			var stylesPart = workBook.AddNewPart<WorkbookStylesPart>();
			stylesPart.Stylesheet = new Stylesheet()
			{

				//default font
				Fonts = new Fonts(new Font(
				new FontSize() { Val = 11D },
				new Color() { Theme = 1U },
				new FontName() { Val = "Calibri" },
				new FontFamily() { Val = 2 },
				new FontScheme() { Val = FontSchemeValues.Minor }))
				{ Count = 1 },

				//default fill
				Fills = new Fills(new Fill(
				new PatternFill() { PatternType = PatternValues.None }))
				{ Count = 1 },

				//default border
				Borders = new Borders(new Border(
				new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder()))
				{ Count = 1 },

				//default cell style
				CellStyleFormats = new CellStyleFormats(new CellFormat
				{
					FillId = 0,
					ApplyFill = false
				})
				{ Count = 1 },

				//this is the default format
				CellFormats = new CellFormats(new CellFormat()
				{
					FontId = 0,
					FillId = 0,
					BorderId = 0,
					FormatId = 0,
					ApplyNumberFormat = false,
					ApplyFont = true
				})
				{ Count = 3 }
			};

			//this is the currency format
			stylesPart.Stylesheet.CellFormats.Append(new CellFormat()
			{
				FontId = 0,
				FillId = 0,
				BorderId = 0,
				FormatId = 0,
				NumberFormatId = 44,
				ApplyNumberFormat = true,
				ApplyFont = true
			});

			//this is the date format
			stylesPart.Stylesheet.CellFormats.Append(new CellFormat()
			{
				FontId = 0,
				FillId = 0,
				BorderId = 0,
				FormatId = 0,
				NumberFormatId = 14,
				ApplyNumberFormat = true,
				ApplyFont = true
			});
		}
	}
}
