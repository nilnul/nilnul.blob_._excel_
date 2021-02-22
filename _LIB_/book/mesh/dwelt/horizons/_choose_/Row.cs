using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.obj._matrix._coord;

namespace nilnul.fs.excel.doc.sheet.dwelt
{

	/// <summary>
	/// including imaginary row (which is not stored with value in the document file.)
	/// </summary>
	static public class _RowX
	{
		//public static Row GetRow(SheetData sheetData, _sheet._coord.Row rowIndex)
		//{
		//	return sheetData.Elements<Row>().Where(r => _sheet._coord.Row.CreateFroRowIndex(r.RowIndex) == rowIndex).FirstOrDefault();
		//}



		// Given a worksheet and a row index, return the row.
		[Obsolete()]

		public static Row GetRow(Worksheet worksheet, uint rowIndex__zeroBased)
		{
			return worksheet.GetFirstChild<SheetData>().
			  Elements<Row>().Where(r => Index(r) == rowIndex__zeroBased).First();
		}

		[Obsolete()]
		public static Row GetRow(Worksheet worksheet, nilnul.Num rowIndex_1base)
		{
			return worksheet.GetFirstChild<SheetData>().
			  Elements<Row>().Where(r => r.RowIndex == rowIndex_1base).First();
		}

		/// <summary>
		/// nullable
		/// </summary>
		/// <param name="sheetData"></param>
		/// <param name="rowIndex"></param>
		/// <returns></returns>
		/// 
		[Obsolete()]
		public static Row GetRow(SheetData sheetData, uint rowIndex)
		{
			return sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).FirstOrDefault();
		}


		[Obsolete()]
		public static Row GetRow(SheetData sheetData, Num rowIndex)
		{
			return sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();

			//throw new NotImplementedException();
		}

		[Obsolete()]
		public static Row GetRow(Worksheet worksheet, int rowIndex_zeroBased)
		{
			return GetRow(worksheet, (uint)rowIndex_zeroBased);
			//throw new NotImplementedException();
		}

		public static Row GetRow(Worksheet worksheet, RowI row)
		{
			return fs.excel.doc.sheet.Dwelt._Create( GetRow(worksheet);
			throw new NotImplementedException();
		}

		/// <summary>
		/// make it zero based
		/// </summary>
		/// <param name="row"></param>
		/// <returns></returns>
		/// 
		[Obsolete(nameof(_sheet._coord.Row))]
		public static nilnul.Num Index(this Row row)
		{
			return row.RowIndex - 1;
		}
	}


}
