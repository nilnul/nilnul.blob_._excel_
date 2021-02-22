using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.obj.matrix;
using nilnul.obj._matrix._coord_._row;

namespace nilnul.blob_.excel.doc.sheet.dwelt.belt.bunch.col
{
	static public class _IntersectedConfinesX
	{
		

		static public IEnumerable<nilnul.obj.matrix.block_.Diag> Enumerate(
			Bunch bunch
			,
			nilnul.obj._mesh._coord.col_.CaptialLetter col
		)
		{

			// get all the cells.
			var cels =nilnul.blob_.excel.doc. sheet.dwelt.belt.bunch.col._CelsX.Enumerate(bunch, col);

			//get all the blocks
			var blocks = cels.Select(
				c => sheet.dwelt.cel._ConfineX.ByMergeCellReference(bunch.belt.dwelt, c)
			);
			return blocks;

		}



		static public nilnul.obj.matrix.block.Set Set(
			Bunch dwelt
			,
			nilnul.obj._mesh._coord.col_.CaptialLetter col
		)
		{

			return new nilnul.obj.matrix.block.Set(
				Enumerate(dwelt, col)
			);

		}






	

	}
}
