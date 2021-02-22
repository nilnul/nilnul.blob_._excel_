using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.fs.excel.doc._sheet._coord_._row;
using nilnul.obj._matrix._coord_._row;

namespace nilnul.fs.excel.doc.sheet.dwelt.row.cels_
{

	static public class _ReifiedX
	{

		

		public static IEnumerable<Cell> Enumerate(Row row
		)
		{

			return row?.Elements<Cell>() ?? new Cell[0];
		}

		public static IEnumerable<Cell> Enumerate(SheetData sheetData,ValI row)
		{
			return Enumerate(dwelt.rows._ChooseX.Get(sheetData, row));


			//throw new NotImplementedException();
		}

		public static IEnumerable<Cell> Enumerate(Worksheet worksheet, _sheet._coord_._row.Val row)
		{
			return Enumerate(dwelt.rows._ChooseX.Get(worksheet, row));


			//throw new NotImplementedException();
		}

		
	}
}
