using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.book.meshs.choose_
{
	static public class BySheet
	{



		public static WorksheetPart GetWorksheetPart(SpreadsheetDocument doc, Sheet sheet)
		{
			string relationshipId = sheet.Id.Value;

			WorksheetPart worksheetPart = (WorksheetPart)
				 doc.WorkbookPart.GetPartById(relationshipId);
			return worksheetPart;

			//throw new NotImplementedException();
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="workbook"> workbook is the root element of workbookPart</param>
		/// <param name="sheet"></param>
		/// <returns></returns>
		/// 
		[Obsolete("workbook is the top element of workbookPart")]
		public static WorksheetPart GetWorksheetPart(Workbook workbook, Sheet sheet)
		{
			string relationshipId = sheet.Id.Value;

			WorksheetPart worksheetPart = (WorksheetPart)
				 workbook.WorkbookPart.GetPartById(sheet.Id);
			return worksheetPart;

			//throw new NotImplementedException();
		}



		public static Worksheet GetWorksheet(SpreadsheetDocument doc, Sheet sheet)
		{
			return GetWorksheetPart(doc, sheet).Worksheet;

			//throw new NotImplementedException();
		}

		[Obsolete("workbook is the top element of workbookPart")]
		public static Worksheet GetWorksheet(Workbook doc, Sheet sheet)
		{
			return GetWorksheetPart(doc, sheet).Worksheet;

			//throw new NotImplementedException();
		}
		public static Worksheet GetWorksheet(WorkbookPart doc, Sheet sheet)
		{
			return GetWorksheet(doc.Workbook, sheet);

			//throw new NotImplementedException();
		}


	}
}
