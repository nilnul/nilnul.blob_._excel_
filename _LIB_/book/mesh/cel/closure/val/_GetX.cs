using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet.dwelt.row.cel.closure.val
{
	public class _GetX
	{
		static public object GetVal(
			WorkbookPart workbookPart
			,
			Worksheet worksheet
			,
			nilnul.obj._matrix.CoordI cel
		) {

			return dwelt.closure.val._GetX.Get(workbookPart,worksheet, 
				nilnul.fs.excel.doc.sheet.dwelt.cel.to_._ClosureX.ByMergeCellReference(worksheet,  cel)
			);



		}
	}
}
