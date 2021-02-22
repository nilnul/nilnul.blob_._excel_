using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.num.duo;
using nilnul.fs.excel.doc._sheet._coord_._col;
using nilnul.fs.excel.doc._sheet._coord_._row;
using nilnul.fs.excel.doc._sheet;
using nilnul.obj._matrix;

namespace nilnul.fs.excel.doc.sheet.dwelt
{
	[Obsolete(nameof(nilnul.blob_.excel.doc.sheet._CelX))]
	public class _CelX
	{

		public static Cell Get(Worksheet worksheet, nilnul.fs.excel.doc._sheet.Coord coord)
		{

			return dwelt.row.cels._ChooseX.ByCoord(worksheet, coord);

			
		}

		public static Cell Get(Worksheet worksheet, _sheet._coord_._row.Val row, _sheet._coord_._col.Val col)
		{
			return Get(worksheet, new _sheet.Coord(row,col));
			//throw new NotImplementedException();
		}

		public static Cell Get(SheetData worksheet, _sheet.Coord coord)
		{

			return dwelt.row.cels._ChooseX.ByCoord(worksheet, coord);

			//throw new NotImplementedException();
		}

		public static Cell Get(SheetData sheetData, obj._matrix.CoordI coord)
		{
			return dwelt.row.cels._ChooseX.ByCoord(sheetData, coord);


			throw new NotImplementedException();
		}
	}
}
