using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheets.choose_
{
	public class ByName
	{
		public static Sheet GetSheet(IEnumerable<Sheet> sheets, string name)
		{
			return sheets.FirstOrDefault(s => string.Compare(s.Name, name, true) == 0);
		}

		public static WorksheetPart
		 GetWorksheetPart(SpreadsheetDocument document,
		 string sheetName)
		{
			IEnumerable<Sheet> sheets =
			   document.WorkbookPart.Workbook.GetFirstChild<Sheets>().
			   Elements<Sheet>().Where(s => s.Name == sheetName);

			if (sheets.Count() == 0)
			{
				// The specified worksheet does not exist.

				return null;
			}

			string relationshipId = sheets.First().Id.Value;
			WorksheetPart worksheetPart = (WorksheetPart)
				 document.WorkbookPart.GetPartById(relationshipId);
			return worksheetPart;

		}

		public static WorksheetPart GetWorksheetPart(WorkbookPart workbookPart, string name)
		{

			IEnumerable<Sheet> sheets = workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();

			string relationshipId = sheets.FirstOrDefault(s => string.Compare(s.Name, "Names", true) == 0).Id;

			var id = sheets.FirstOrDefault(s => string.Compare(s.Name, name, true) == 0).Id;

			return (WorksheetPart)workbookPart.GetPartById(id);


		}

	}
}
