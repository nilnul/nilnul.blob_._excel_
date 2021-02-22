using nilnul.obj.mesh._bloc;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.doc.sheet.block
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
			Sheet sheet
			,
			nilnul.blob_.excel.doc.sheet._bloc.Bounds bounds

		)
		{


			var r = new System.Data.DataTable();
			var sheetFirst = sheet.worksheet;//  nilnul.blob_.excel.doc.sheet_._FirstX.GetFirst(file);

			var firstRow_ord = bounds.rowBound.lower;
			var firstRow = new nilnul.blob_.excel.doc.sheet.bloc_._row.Bounds(bounds.colBound, firstRow_ord);




			//var cels = nilnul.blob_.excel.doc.sheet.dwelt.row._CelsX.Enumerate(

			//			sheetFirst,
			//			bounds.rowBound.lower
			//		//row
			//		);


			r.Columns.AddRange(
				nilnul.blob_.excel.doc.sheet.bloc_.row.Vals.GetTxt(sheet, firstRow).Select(x=> new DataColumn(x)).ToArray()



			);


			var rowBegin = bounds.rowBound.lower.toNum() +1; // the previous of data.


			var rowEnd = bounds.rowBound.upper.toNum();


			for (var  row = rowBegin;row<= rowEnd; row = row+1)
			{

				var cels1 = nilnul.blob_.excel.doc.sheet.bloc_.row.Vals.Enumerate(
					sheet,
					new nilnul.obj.mesh._bloc.bounds_.Row1<num.ord_.oneBased_.bijective_.UpperLetter,  num.Ord2>(
						bounds.colBound, 
						new nilnul.num.Ord2(
							row
						)
					)
				).ToArray();

				

				r.Rows.Add(
					cels1
				);
			}
			return r;
		}

		public static DataTable GetDataTbl(Sheet sheet, Bounds bounds)
		{
			return GetDataTbl(sheet, new _bloc.Bounds(bounds) );
		}
	}
}

