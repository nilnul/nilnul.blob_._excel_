using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.book.mesh.dwelt.horizon.cels_
{

	static public class _NonblankX
	{
		static public nilnul.txt.Set TxtSet(
			book.mesh.dwelt.Horizon level

		)
		{
			var celsReified = blob_.excel.book.mesh.dwelt.horizon.cels_._ReifiedX.Cells(level);

			return new txt.Set(
				celsReified.Select(c =>

				mesh.cel.val_.typed._GetX.Obj(level.boxed.mesh, c).ToString()
				 )
			);

		}


		public static nilnul.txt.Set TxtSet(Dwelt dwelt, Row current)
		{
			return TxtSet(
				new Horizon(dwelt, current)
			);
		}
	}
}
