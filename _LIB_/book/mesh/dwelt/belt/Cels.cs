using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.obj._matrix._coord_._row;

namespace nilnul.blob_.excel.doc.sheet.dwelt.belt
{
	/// <summary>
	/// return the coord
	/// </summary>
	static public class _CelsX
	{
		
		public static IEnumerable<nilnul.obj._mesh.CoordI> Enumerate(

			sheet.dwelt.Belt belt
			
		)
		{

			return nilnul.obj._mesh.range.Cels.Enumerate(
				 new nilnul.obj._mesh.Range(
					 num.Bound.CreateClose( belt.rows.lower.toNum(),belt.rows.upper.toNum())

				,
					  num.Bound.CreateClose( belt.cols.lower,belt.cols.upper)
				)
			);

		
			

		}



		


	}
}
