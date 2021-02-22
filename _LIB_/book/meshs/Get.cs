using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheets
{
	public class _GetX
	{

		[Obsolete(nameof(GetSheets_workbookSheets))]
		public static Sheets GetSheets(SpreadsheetDocument spreadSheetDocument)
		{


			WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
			return spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();//.Elements<Sheet>();

		}

		public static Sheets GetSheets_workbookSheets(SpreadsheetDocument spreadSheetDocument)
		{


			WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
			return spreadSheetDocument.WorkbookPart.Workbook.Sheets;//.Elements<Sheet>();

		}


		public static IEnumerable<Sheet> GetSheetEnumerable(SpreadsheetDocument spreadSheetDocument)
		{
			return GetSheets(spreadSheetDocument).Elements<Sheet>();

		}
		static public IEnumerable< Sheet> GetSheetEnumerable_byWorkbookDescendants(WorkbookPart  workbookPart)
		{

			return workbookPart.Workbook.Descendants<Sheet>();

			

		}
		[Obsolete(nameof(GetSheetEnumerable_byWorkbookSheets))]
		static public IEnumerable<Sheet> GetSheetEnumerable(WorkbookPart workbookPart)
		{
			return workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
		}

		static public IEnumerable<Sheet> GetSheetEnumerable_byWorkbookSheets(WorkbookPart wbPart)
		{
			return wbPart.Workbook.Sheets.Elements<Sheet>();

		}



		static public IEnumerable< WorksheetPart> GetWorksheetParts_ordReversed(WorkbookPart  workbookPart)
		{
			/*
			think this is caused by the fact that you have only one sheet whereas Excel has three.I'm not certain but I think the sheets are returned in reverse order so you should change the line:
WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

			to
			WorksheetPart worksheetPart = workbookPart.WorksheetParts.Last();

			It might be safer to search for the WorksheetPart if you can identify it by the sheet name.You need to find the Sheet first then use the Id of that to find the SheetPart:
			*/

			return workbookPart.WorksheetParts;
		}





		






	}
}
