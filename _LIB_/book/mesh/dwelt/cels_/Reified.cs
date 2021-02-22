using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet.dwelt.cels
{
	public class _ReifiedX
	{
		static public IEnumerable<Cell> Enumerate(WorksheetPart wsPart)
		{
			/// cell must be under Row; or so?
			return wsPart.Worksheet.Descendants<Cell>();//.Where(c => c.CellReference == addressName).FirstOrDefault();

		}

		
	}
}
