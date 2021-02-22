using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using nilnul.objs.be_;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.obj._matrix._coord_._row;
using nilnul.num.ord_;
using nilnul.obj.matrix;

namespace nilnul.blob_.excel.doc.sheet.dwelt.belt
{
	/// <summary>
	/// blocs of a belt.
	/// 
	/// </summary>
	static public class _ConfinesX
	{



		/// <summary>
		/// initially, a belt is ending before the row; Now for this row, we find its blocs (mergedCels).
		/// during the finding, we recursively merging lastly-appended row's blocs until no new row is appended.
		/// </summary>
		/// <param name="worksheet"></param>
		/// <param name="_row_first"></param>
		/// <returns></returns>


		public static IEnumerable<nilnul.obj.matrix.BlockI> _Enumerate(Worksheet worksheet,  ValI _row_first)
		{

			var dim = dwelt._DiagX.Get(worksheet);



			var blocks = dwelt.row._IntersectConfinesX.GetSet(worksheet, _row_first);

			var endRow = nilnul.obj.matrix.block.str_.started.EndRow._Get(
				blocks
			);

			if (nilnul.obj._matrix._coord_._row.val.Eq.Singleton.Equals(_row_first, endRow))
			{
				return blocks; //case closed, as the row is not convoluted with any other row.
			}

			else
			{

				return
					blocks.Union(
						_Enumerate(
							worksheet,
							 new nilnul.fs.excel.doc. _sheet._coord_._row.Val(
								endRow	//open for further implication.
							)
						)
					)
				;
			}
		}

		private static IEnumerable<BlockI> _Enumerate(Worksheet worksheet, OneBased item1)
		{
			return _Enumerate(worksheet, new nilnul.obj._matrix._coord_._row.Val(item1.toNum()));
		}

		public static IEnumerable<nilnul.obj.matrix.BlockI> Enumerate(Belt belt ) {
			return _Enumerate( belt.dwelt.sheet.worksheet, belt.rows.en.Item1);
		}


		public static nilnul.obj.matrix.block.Set GetSet(Belt belt)
		{
			return new nilnul.obj.matrix.block.Set(
				Enumerate(
					 belt
				 )
			);

			//throw new NotImplementedException();
		}

		
	}
}
