using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using nilnul.objs.be_;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.obj._matrix._coord_._row;

namespace nilnul.fs.excel.doc.sheet.dwelt.row_.mergedNoPre
{
	/// <summary>
	/// 
	/// 
	/// </summary>
	static public class _BeltX
	{

		/// <summary>
		/// get the recleared row 's belt. 
		/// </summary>
		/// <param name="worksheet"></param>
		/// <param name="row"></param>
		/// <returns></returns>
		

		public static IEnumerable<nilnul.obj.matrix.BlockI> Enumerate(Worksheet worksheet,  ValI row)
		{

			var dim = dwelt._AsDiagX.Get(worksheet);



			var blocks = dwelt.row._GetIntersectBlocksX.GetSet(worksheet, row);

			var endRow = nilnul.obj.matrix.block.str_.started.EndRow._Get(
				blocks
			);

			if (nilnul.obj._matrix._coord_._row.val.Eq.Singleton.Equals(row, endRow))
			{
				return blocks; //case closed
			}

			else
			{

				return
					blocks.Union(
						Enumerate(
							worksheet,
							 new _sheet._coord_._row.Val(
								endRow	//open for further implication.
							)
						)
					)
				;
			}
		}

		public static nilnul.obj.matrix.block.Set GetSet(Worksheet worksheet, Val row)
		{
			return new nilnul.obj.matrix.block.Set(
				 Enumerate(
					 worksheet, row
					 )
			);

			//throw new NotImplementedException();
		}

		static public nilnul.obj.matrix.block.Set GetSet(
			Worksheet worksheet
			,
			_sheet._coord_._row.Val row
		)
		{
			return new nilnul.obj.matrix.block.Set(
				 Enumerate(
					 worksheet, row
					 )
			);
		}
	}
}
