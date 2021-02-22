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

namespace nilnul.blob_.excel.book.mesh.row.cels_
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
			return Enumerate(rows_.dwelt._ChooseX.Get(sheetData, row));


			//throw new NotImplementedException();
		}

		public static IEnumerable<Cell> Enumerate(Worksheet worksheet, nilnul.num.ord_.OneBased row)
		{
			return Enumerate(rows_.dwelt._ChooseX.Get(worksheet, row));


			//throw new NotImplementedException();
		}

		public static IEnumerable<Cell> Enumerate<TRow>(SheetData sheetData, TRow row) where TRow : OrdI2
		{
			return Enumerate(rows_.dwelt._ChooseX.Get(sheetData, row));

		}
	}
}
