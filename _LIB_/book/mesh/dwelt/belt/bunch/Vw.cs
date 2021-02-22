using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.doc.sheet.dwelt.belt.bunch
{
	static public class _CelsX
	{
		static public IEnumerable<nilnul.blob_.excel.doc._sheet.Coord> Coords(Bunch bunch ) {
			var colBound = bunch.cols;
			var rowBound = bunch.belt.rows;

			for (var  i = rowBound.lower; nilnul.num.ord.comparer.Re.Singleton.le(i,  rowBound.upper) ; i++  )
			{
				for (var j = colBound.lower; j <=  colBound.upper; j++)
				{
					yield return new  blob_.excel.doc._sheet.Coord(new nilnul.obj._mesh._coord.Row(  i.toNum()) , new obj._mesh._coord.Col(  j.toNum()));

				}
			}
		}
	}
}
