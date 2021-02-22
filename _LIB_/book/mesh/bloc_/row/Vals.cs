using DocumentFormat.OpenXml.Spreadsheet;
using nilnul.num;
using nilnul.num.ord_.oneBased_.bijective_;
using nilnul.obj.mesh._bloc.bounds_;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.doc.sheet.bloc_.row
{
	/// <summary>
	/// enumerate the values of each cel. (for unpopulated cell, the val is null)
	/// </summary>
	static public class Vals
	{
		public static IEnumerable<object> Enumerate(

			Sheet dwelt
			,
			nilnul.blob_.excel.doc.sheet.bloc_._row.Bounds row
		)
		{
			return nilnul.obj.mesh._bloc.bounds_.row.cels._EnumerateX.Enumerate<num.ord_.oneBased_.bijective_.UpperLetter, num.ord_.OneBased>(
			   row
			  

		   ).Select(x => excel.book.mesh.cel.val_.typed._GetX.GetObj(dwelt, x));

		}

		//public static IEnumerable<object> Enumerate(

		//	Sheet dwelt
		//	,
		//	nilnul.blob_.excel.doc.sheet.bloc_._row.Bounds row
		//)
		//{
		//	return nilnul.obj.mesh._bloc.bounds_.row.cels._EnumerateX.Enumerate<num.ord_.oneBased_.bijective_.UpperLetter, num.ord_.OneBased>(
		//	   row
			  

		//   ).Select(x => excel.doc.sheet.cel.val_.typed._GetX.GetObj(dwelt, x));

		//}
		public static IEnumerable<object> Enumerate(Sheet sheetFirst, Row1<UpperLetter, Ord2> row)
		{

			return nilnul.obj.mesh._bloc.bounds_.row.cels._EnumerateX.Enumerate<num.ord_.oneBased_.bijective_.UpperLetter, num.Ord2>(
			   row
			  

		   ).Select(x => excel.book.mesh.cel.val_.typed._GetX.GetObj(sheetFirst, x));


		}

		public static IEnumerable<string> GetTxt(

			Sheet dwelt
			,
			nilnul.blob_.excel.doc.sheet.bloc_._row.Bounds row
		)
		{
			return nilnul.obj.mesh._bloc.bounds_.row.cels._EnumerateX.Enumerate<num.ord_.oneBased_.bijective_.UpperLetter, num.ord_.OneBased>(
			   row
			  

		   ).Select(x => excel.doc.sheet.cel.val.asTxt._VwX.GetTxt(dwelt, x));

		}

	}
}
