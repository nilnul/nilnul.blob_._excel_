using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.num.duo;
using nilnul.fs.excel.doc._sheet;
using nilnul.obj._matrix;

namespace nilnul.fs.excel.doc.sheet.dwelt.row.cels
{
	public class _ChooseX
	{
		// Given a worksheet, a column name, and a row index, 
		// gets the cell at the specified column and 
		public static Cell ByCoord(
			Worksheet worksheet
			,
			 _sheet.Coord coord

			)
		{
	
			DocumentFormat.OpenXml.Spreadsheet.Row row = rows._ChooseX.Get(worksheet, coord.row);


			if (row == null)
				return null;

			var cels = cels_._ReifiedX.Enumerate(worksheet, coord.row);

			return cels.Where(
				c => nilnul.obj._matrix.coord.Eq.Singleton.Equals(
					_sheet.Coord.CreateFroCell(c)
					,
					coord
				)
			).FirstOrDefault();
		}

		internal static Cell ByCoord(SheetData sheetData, _sheet.Coord coord)
		{

			DocumentFormat.OpenXml.Spreadsheet.Row row = rows._ChooseX.Get(sheetData, coord.row);


			if (row == null)
				return null;

			

			var cels = cels_._ReifiedX.Enumerate(sheetData, coord.row);

			return cels.Where(
				c => nilnul.obj._matrix.coord.Eq.Singleton.Equals(
					_sheet.Coord.CreateFroCell(c)
					,
					coord
				)
			).FirstOrDefault();

			//throw new NotImplementedException();
		}

		internal static Cell ByCoord(SheetData sheetData, obj._matrix.CoordI coord)
		{

			DocumentFormat.OpenXml.Spreadsheet.Row row = rows._ChooseX.Get(sheetData, coord.row);


			if (row == null)
				return null;

			

			var cels = cels_._ReifiedX.Enumerate(sheetData, coord.row);

			return cels.Where(
				c => nilnul.obj._matrix.coord.Eq.Singleton.Equals(
					_sheet.Coord.CreateFroCell(c)
					,
					coord
				)
			).FirstOrDefault();



			throw new NotImplementedException();
		}
	}
}
