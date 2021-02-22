using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.obj.matrix;
using nilnul.obj._matrix._coord_._row;

namespace nilnul.blob_.excel.doc.sheet.dwelt.belt.col
{
	static public class _IntersectedConfinesX
	{
		

		static internal IEnumerable<nilnul.obj.matrix.block_.Diag> _Enumerate(
			Belt belt
			,
			nilnul.obj._mesh._coord.col_.CaptialLetter col
		)
		{

			// get  the cells along the col
			var cels =nilnul.blob_.excel.doc. sheet.dwelt.belt.col._CelsX.Enumerate(belt, col);

			//get all the blocks
			var blocks = cels.Select(
				c => sheet.dwelt.cel._ConfineX.ByMergeCellReference(belt.dwelt, c)
			);
			return blocks;

		}



		static public nilnul.obj.matrix.block.Set Set(
			Belt dwelt
			,
			nilnul.obj._mesh._coord.col_.CaptialLetter col
		)
		{

			return new nilnul.obj.matrix.block.Set(
				_Enumerate(dwelt, col)
			);

		}






	

	}
}
