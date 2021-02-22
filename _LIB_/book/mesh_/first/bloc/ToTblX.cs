using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.doc.sheet_.first.block
{
	static public class ToTblX
	{
		/// <summary>
		/// first line is cols
		/// </summary>
		/// <param name="file"></param>
		/// <param name="bounds"></param>
		/// <returns></returns>
		static public DataTable GetDataTbl(
			string file
			,
			nilnul.obj.mesh._bloc.Bounds bounds

		)
		{
			return nilnul.blob_.excel.doc.sheet.block.ToTblX.GetDataTbl(
				nilnul.blob_.excel.doc.sheet_._FirstX.GetFirst(file)
				,
				bounds
			);

				
			

		}
	}
}
