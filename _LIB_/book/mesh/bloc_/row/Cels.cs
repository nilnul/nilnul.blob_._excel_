using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.doc.sheet.block_.row
{
	/// <summary>
	/// enumerate the cels. (for unpopulated cell, null is assumed)
	/// </summary>
	static public class Cels
	{
		public static IEnumerable<nilnul.blob_.excel.doc._sheet.Coord> Enumerate(

	Sheet dwelt
	,
	nilnul.blob_.excel.doc.sheet.bloc_._row.Bounds row
)
		{
			

			return nilnul.obj.mesh._bloc.bounds_.row.cels._EnumerateX.Enumerate(
				row
				,
				new nilnul.obj._matrix._coord_._row.Val(
					row.toNum()
				)

			).Select(x => new nilnul.blob_.excel.doc._sheet.Coord(x));

		}
	}
}
