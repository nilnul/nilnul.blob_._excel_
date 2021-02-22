using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet.dwelt.closures
{
	/// <summary>
	/// divide the closures into belts such that:
	/// 1) each belt cannot be divided into smaller belt.
	/// </summary>
	/// 
	[Obsolete(nameof(Belts))]
	static public class _MinBeltSX
	{
		static public IEnumerable<_sheet._coord_._row.val.Range> Enumerate(Worksheet worksheet)
		{

			var rowsContiguous = dwelt._rows._GetX.RowIndexS(worksheet);


			var enumerator = rowsContiguous.GetEnumerator();

			//var started = enumerator.MoveNext();

			//	uint rowIndex_1based;
			//uint rowIndex_0based=0;

			if (enumerator.MoveNext())
			{
				_sheet._coord_._row.Val startRowIndex= new _sheet._coord_._row.Val(enumerator.Current); //the sheet may start at 3, for example


				do
				{

					//var currentRow = enumerator.Current;

					//if (currentRow == null) //row with no data 
					//{
					//	yield return _sheet._coord_._row.val.Range.CreateSingleton(startRowIndex);


					//}
					//else
					//{

					var boundingBelt = row_.mergedNoPre.Belt_retRange._Get(worksheet, startRowIndex);

					yield return boundingBelt;

					startRowIndex = _sheet._coord_._row.val._GrowX.Grow(
						startRowIndex
						,
						boundingBelt.count
					);

					nilnul.obj.seq.enumerator._MoveNextX.MoveNext(
						enumerator,
						(int)boundingBelt.count.toBigint() - 1 //we 'll move one row more later.
					);
					//}
				} while (enumerator.MoveNext());
			}

		}
		/// <summary>
		/// 
		/// </summary>
		/// <param name="worksheetPart"></param>
		/// <returns></returns>
		/// <remarks>
		/// use worksheet
		/// </remarks>
		static public IEnumerable<_sheet._coord_._row.val.Range> Enumerate(WorksheetPart worksheetPart)
		{
			return Enumerate(worksheetPart.Worksheet);

		}



	}
}
