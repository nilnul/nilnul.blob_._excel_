using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.obj._matrix._coord_._row;

namespace nilnul.blob_.excel.doc.sheet.dwelt.belt.col
{
	/// <summary>
	/// return the coord
	/// </summary>
	static public class _CelsX
	{
		
		public static IEnumerable<nilnul.obj._mesh.Coord> Enumerate(

			sheet.dwelt.Belt belt
			,
			nilnul.obj._mesh._coord.col_.CaptialLetter col
		)
		{
			var sub = new nilnul.obj._mesh.Range(
				 num.Bound.CreateClose( belt.rows.lower.toNum(),belt.rows.upper.toNum())

			,
				  num.Bound.CreateClose( belt.cols.lower,belt.cols.upper)
			);

			for (var row  = sub.vertical.ed.lower.mark; row <= sub.vertical.ed.upper.mark; row=row+1)
			{

				
					yield return  obj._mesh.Coord.FroRowCol( row, col);
			}
			
		}



		


	}
}
