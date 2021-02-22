using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheets.choose_
{
	static public class _FirstX
	{

		

		public static Sheet GetSheet(SpreadsheetDocument document)
		{
			return sheets._GetX.GetSheetEnumerable(document).First();
		}

		static public Worksheet GetWorksheet(
				SpreadsheetDocument doc
	
			) {
			return GetWorksheet(doc.WorkbookPart);
		}


		public static WorksheetPart GetWorksheetPart(SpreadsheetDocument document)
		{

			Sheet theSheet = document.WorkbookPart.Workbook.Descendants<Sheet>().First();

			return (WorksheetPart)(document.WorkbookPart.GetPartById(theSheet.Id)); ;
		}


		static WorksheetPart GetWorksheetPart(WorkbookPart wbPart)
		{
			return (wbPart.GetPartById(

				sheets._GetX.GetSheetEnumerable_byWorkbookSheets(wbPart).First().Id.Value
			) as WorksheetPart)
				;//.Worksheet.SheetDimension.Reference.Value;

		}
		static Worksheet GetWorksheet(WorkbookPart wbPart)
		{
			return GetWorksheetPart(wbPart).Worksheet
				;//.Worksheet.SheetDimension.Reference.Value;

		}


		

	}
}
