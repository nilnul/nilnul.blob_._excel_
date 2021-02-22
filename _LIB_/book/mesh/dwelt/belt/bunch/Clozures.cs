using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.doc.sheet.dwelt.belt.bunch
{
	static public class _ConfinesX
	{
		public static IEnumerable<nilnul.obj.matrix.BlockI> _Enumerate(
			Bunch bunch
			, 
			nilnul.obj._mesh._coord.col_.CaptialLetter _col_first)
		{
			var blocks = belt.bunch.col._IntersectedConfinesX.Set(bunch, _col_first);

			var endCol = nilnul.obj._mesh.range.set_.sed.EndCol._Get(
				blocks
			);

			if (nilnul.obj._mesh._coord.col.Eq.Singleton.Equals(_col_first, endCol))
			{
				return blocks; //case closed
			}

			else
			{

				return
					blocks.Union(
						_Enumerate(
							bunch,
							 new nilnul.obj._mesh._coord.col_.CaptialLetter(
								endCol  //open for further implication.
							)
						)
					)
				;
			}
		}


		/// <summary>
		/// repeated elements
		/// </summary>
		/// <param name="worksheet"></param>
		/// <param name="row"></param>
		/// <returns></returns>
		public static IEnumerable<nilnul.obj.matrix.BlockI> GetIntersectBlocks( Bunch row)
		{
			return _Enumerate(
				row,
				new obj._mesh._coord.col_.CaptialLetter(
					row.cols.First()
				)
			);
		}
	}
}
