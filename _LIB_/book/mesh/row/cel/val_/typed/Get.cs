using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.book.mesh.row.col.val_.typed
{
	static public class _GetX
	{
		public static object Obj_colOneBased(mesh.Major major, int col)
		{

			return excel.book.mesh.cel.val_.typed._GetX.Obj(
				new excel.book.mesh.Cel (
					major.mesh 
				,
				excel.book.mesh.row.cels_.dwelt._ChooseX.Cel0nul_ofColOneBased(major.row, col)
				)
			);
		}

		public static object Obj_colNilBased(mesh.Major major, int col)
		{

			return excel.book.mesh.cel.val_.typed._GetX.Obj(
				new excel.book.mesh.Cel (
					major.mesh 
				,
				excel.book.mesh.row.cels_.dwelt._ChooseX.Cel0nul(major.row, col)
				)
			);
		}


	}
}
