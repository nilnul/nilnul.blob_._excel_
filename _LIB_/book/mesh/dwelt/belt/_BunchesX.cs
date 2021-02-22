using DocumentFormat.OpenXml.Spreadsheet;
using nilnul.fs.excel.doc.sheet.dwelt.closures;
using nilnul.fs.excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.doc.sheet.dwelt.belt
{
	/// <summary>
	/// get the colBounds that are:
	/// 1) no closure across 2 or more colBounds
	/// 2) the colBounds cannot bre divided into more bunches.
	/// </summary>
	static public class _BunchesX
	{

		static public IEnumerable<blob_.excel.doc._sheet._coord.col.Bound> HorizontalBoundS(Belt belt)
		{

			var enumerator =belt.cols.GetEnumerator();

			


			while (enumerator.MoveNext())
			{
				var bunch = col_.mergedNoPre.bunch.range.bound_._HorizontalX._Get(
					belt.dwelt.sheet.worksheet
					,
					enumerator.Current
				);

				yield return bunch;

				nilnul.obj.seq.enumerator._MoveNextX.MoveNext(
					enumerator,
					(int)bunch.count.toBigint() - 1 //we 'll move one row more later.
				);
			}
		}

		
	}
}
