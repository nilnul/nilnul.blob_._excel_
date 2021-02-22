using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.book.mesh
{
	static public class DweltX
	{
		static public  SheetData GetSheetData(Worksheet worksheet)
		{
			return worksheet.GetFirstChild<SheetData>();

		}

		static public  SheetData GetSheetData(WorksheetPart worksheetPart)
		{
			return GetSheetData(worksheetPart.Worksheet); 

		}





		/// <summary>
		/// 
		/// </summary>
		/// <param name="sheet">in workbook element's sheets element</param>
		/// <returns></returns>
		public static SheetData GetSheetData(SpreadsheetDocument doc, DocumentFormat.OpenXml.Spreadsheet.Sheet sheet)
		{
			return GetSheetData(
				nilnul.fs.excel.doc.sheets.choose_.BySheet.GetWorksheet(doc,sheet)
			);

			//throw new NotImplementedException();
		}

		public static SheetData GetSheetData(Workbook doc, DocumentFormat.OpenXml.Spreadsheet.Sheet sheet)
		{
			return GetSheetData(
				nilnul.fs.excel.doc.sheets.choose_.BySheet.GetWorksheet(doc,sheet)
			);

			//throw new NotImplementedException();
		}



	}
}
