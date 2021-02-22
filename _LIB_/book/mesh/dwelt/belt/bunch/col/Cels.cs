using System.Collections.Generic;

namespace nilnul.blob_.excel.doc.sheet.dwelt.belt.bunch.col
{
	/// <summary>
	/// return the coord
	/// </summary>
	public static class _CelsX
	{

		public static IEnumerable<nilnul.obj._mesh.Coord> Enumerate(

			sheet.dwelt.belt.Bunch bunch
			,
			nilnul.obj._mesh._coord.col_.CaptialLetter col
		)
		{
			obj._mesh.Range range = new nilnul.obj._mesh.Range(
				 num.Bound.CreateClose(bunch.belt.rows.lower.toNum(), bunch.belt.rows.upper.toNum())

			,
				  num.Bound.CreateClose(bunch.cols.lower, bunch.cols.upper)
			);

			for (Num row = range.vertical.ed.lower.mark; row <= range.vertical.ed.upper.mark; row = row + 1)
			{


				yield return obj._mesh.Coord.FroRowCol(row, col);
			}

		}






	}
}
