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

namespace nilnul.blob_.excel.doc.sheet
{
	public class _CelX
	{
		// Given a worksheet, a column name, and a row index, 
		// gets the cell at the specified column and 
		public static Cell ByCoord(
			Worksheet worksheet
			,
			sheet._cel.Coord coord
			// _sheet.Coord coord

			)
		{
	
			DocumentFormat.OpenXml.Spreadsheet.Row row = sheet.dwelt.rows._ChooseX.Get(worksheet, coord.row);


			if (row == null)
				return null;

			var cels = sheet.dwelt.row.cels_._ReifiedX.Enumerate(worksheet, coord.row);

			return cels.Where(
				c => nilnul.obj.mesh._cel.coord.Eq<nilnul.num.ord_.oneBased_.bijective_.UpperLetter,nilnul.num.ord_.OneBased>.Singleton.Equals(
					sheet.cel._CoordX.CreateFroCell(c)
					,
					coord
				)
			).FirstOrDefault();
		}

		public static Cell ByCoord(SheetData sheetData, sheet._cel.Coord coord)
		{

			DocumentFormat.OpenXml.Spreadsheet.Row row = sheet.dwelt.rows._ChooseX.Get(sheetData, coord.row);

			//	rows._ChooseX.Get(sheetData, coord.row);


			if (row == null)
				return null;

			

			var cels = sheet.dwelt.row.cels_._ReifiedX.Enumerate(sheetData, coord.row);

			return cels.Where(
				c => nilnul.blob_.excel.doc.sheet._cel.coord.Eq.Singleton.Equals(

					sheet.cel._CoordX.CreateFroCell(c)
					,
					coord
				)
			).FirstOrDefault();

			//throw new NotImplementedException();
		}

		public static Cell ByCoord<TCol,TRow>(SheetData sheetData, obj.mesh._cel.CoordI<TCol,TRow> coord)
			where TCol: nilnul.num.OrdI2
			where TRow: nilnul.num.OrdI2

		{

			DocumentFormat.OpenXml.Spreadsheet.Row row = sheet.dwelt.rows._ChooseX.Get(sheetData, coord.row);


			if (row == null)
				return null;



			var cels = sheet.dwelt.row.cels_._ReifiedX.Enumerate(sheetData, coord.row);

			return cels.Where(
				c => nilnul.num.ord.co.Eq.Singleton.Equate(
					(nilnul.num.ord.CoI)
					nilnul.blob_.excel.doc.sheet._cel.Coord.CreateFroCell(c)
					,
					(nilnul.num.ord.CoI)

					coord
				)
			).FirstOrDefault();



			throw new NotImplementedException();
		}
	}
}
