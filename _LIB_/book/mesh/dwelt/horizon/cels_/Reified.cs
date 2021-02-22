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
using nilnul.num;

namespace nilnul.blob_.excel.book.mesh.dwelt.horizon.cels_
{

	static public class _ReifiedX
	{

		

		public static IEnumerable<Cell> Enumerate(Row row
		)
		{

			return row?.Elements<Cell>() ?? new Cell[0];
		}

		public static IEnumerable<Cell> Enumerate(SheetData sheetData,nilnul.num.ord_.OneBased row)
		{
			return Enumerate(dwelt.horizons._ChooseX.Get(sheetData, row));


			//throw new NotImplementedException();
		}

		public static IEnumerable<Cell> Enumerate(Worksheet worksheet, nilnul.num.ord_.OneBased row)
		{
			return Enumerate(dwelt.horizons._ChooseX.Get(worksheet, row));


			//throw new NotImplementedException();
		}

		public static IEnumerable<Cell> Enumerate<TRowOrd>(SheetData sheetData, TRowOrd row) where TRowOrd : OrdI2
		{
			return Enumerate(dwelt.horizons._ChooseX.Get(sheetData, row));

		}

		public static IEnumerable<Cell> Cells(Horizon level)
		{

			return Enumerate(level.row);
		}
	}
}
