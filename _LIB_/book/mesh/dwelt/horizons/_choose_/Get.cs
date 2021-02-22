using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet.dwelt.rowN
{
	[Obsolete(nameof(_RowX))]
	static public class _GetX
	{
		


		// Given a worksheet and a row index, return the row.
		public static Row Get(Worksheet worksheet, uint rowIndex__zeroBased)
		{
			return worksheet.GetFirstChild<SheetData>().
			  Elements<Row>().Where(r => _RowX.Index( r) == rowIndex__zeroBased).FirstOrDefault();
		}

		public static Row Get(Worksheet worksheet, nilnul.Num rowIndex_1base)
		{
			return worksheet.GetFirstChild<SheetData>().
			  Elements<Row>().Where(r => r.RowIndex == rowIndex_1base).FirstOrDefault();
		}

		/// <summary>
		/// nullable
		/// </summary>
		/// <param name="sheetData"></param>
		/// <param name="rowIndex"></param>
		/// <returns></returns>
		public static Row Get(SheetData sheetData, uint rowIndex)
		{
			return sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).FirstOrDefault();
		}

		public static Row Get(SheetData sheetData, Num rowIndex)
		{
			return sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).FirstOrDefault();

			//throw new NotImplementedException();
		}

		public static Row Get(Worksheet worksheet, int rowIndex_zeroBased)
		{
			return Get(worksheet, (uint)rowIndex_zeroBased);
			//throw new NotImplementedException();
		}

		
	}


}
