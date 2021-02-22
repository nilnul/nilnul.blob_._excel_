using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.doc.sheet.dwelt.cels_.reified
{
	public class _VwX
	{
		static public IEnumerable<Cell> Enumerate(WorksheetPart wsPart)
		{
			return wsPart.Worksheet.Descendants<Cell>();//.Where(c => c.CellReference == addressName).FirstOrDefault();

		}

		
	}
}
