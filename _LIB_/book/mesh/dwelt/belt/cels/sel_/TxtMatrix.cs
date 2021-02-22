using nilnul.fs.excel.doc.sheet.dwelt.closures_;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet.dwelt.closures.belt
{
	/// <summary>
	/// get the txts of all cels
	/// </summary>
	public class _GetX
	{
		/// <summary>
		/// get the txt of each cell of the row span
		/// </summary>
		/// <returns></returns>
		static public string[,] GetTxtMatrix(

			Belt belt
		)
		{

			// get the dimension

			var dim = (belt.rows);

			var cols = belt.cols.ToArray();
			var leastCol = cols.First();

			var rows = belt.rows.ToArray();
			var leastRow = Enumerable.First<obj._matrix._coord_._row.ValI>(rows);

			var val = new string[(int)rows.Length, (int)cols.Length];

			for (int row = 0; row < Enumerable.Count<obj._matrix._coord_._row.ValI>(rows); row++)
			{
				for (int col = 0; col < cols.Count(); col++)
				{
					val[row, col] = dwelt.cel.val._GetX.GetTxt(
						belt.dwelled.workbookPart
						,
						belt.dwelled.worksheet
						,

						 new nilnul.fs.excel.doc._sheet.Coord(
							 cols.ElementAt(col), 
							 rows.ElementAt(row)
						 )
					);


				}
			}
			
			return val;

			throw new NotImplementedException();


		}

		/// <summary>
		/// may be empty for a col that intersect with a block's cell that's not the upperLeft.
		/// </summary>
		/// <param name="belt"></param>
		/// <returns></returns>
		static public IEnumerable<string> GetTxts(Belt belt) {

			return nilnul.obj._MatrixX.Cols(
				GetTxtMatrix(belt)
			).Select(
				col=> nilnul.txts.accumulate_.join_.Empty.Singleton.accumulate(col)
			);

		}

	}
}
