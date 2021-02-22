using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.obj._matrix._coord_._row;

namespace nilnul.blob_.excel.doc.sheet.dwelt.col
{
	/// <summary>
	/// return the coord
	/// </summary>
	static public class _CelsX
	{
		
		public static IEnumerable<nilnul.blob_.excel.doc._sheet.Coord> Enumerate(

			sheet.Dwelt dwelt
			,
			nilnul.obj._mesh._coord.col_.CaptialLetter col
		)
		{

			var sub = nilnul.obj.matrix.BlockIX.ToSub(dwelt);

			return nilnul.obj._mesh.grid.col._CelsX.Enumerate(

				new obj._mesh.Grid(
					new obj._mesh._coord.row.Set(
						sub.rows.Select(x=> new nilnul.obj._mesh._coord.Row( new nilnul.Num( x.toNum())  ))
					)
					,
					new obj._mesh._coord.col.Set(
						sub.cols.Select(y=> new nilnul.obj._mesh._coord.Col( new nilnul.Num( y.toNum()) ))
					)
				)
				,
				col

			).Select(x=> new nilnul.blob_.excel.doc._sheet.Coord(  x  ) );

		}



		


	}
}
