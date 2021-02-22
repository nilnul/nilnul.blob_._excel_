using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheets.choose_
{
	public class ById
	{

		static public WorksheetPart Get(SpreadsheetDocument doc, string sheetRelId ) {
			return (WorksheetPart)doc.WorkbookPart.GetPartById(sheetRelId);

		}
	}
}
