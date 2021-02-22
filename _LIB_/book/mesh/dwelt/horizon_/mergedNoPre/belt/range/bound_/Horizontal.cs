using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using nilnul.obj._matrix._coord_._row;
using nilnul.obj.matrix.block;
using nilnul.objs.be_;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet.dwelt.row_.mergedNoPre.belt.range.bound_
{
	/// <summary>
	/// get the outer rowSpan of this row; this row has no cell merged with previous row.
	/// </summary>
	/// 
	static public class _VerticalX
	{



		static public _sheet._coord_._row.val.Range _Get(
			WorksheetPart worksheetPart
			,
			Row _mergedNoPre
		)
		{
			var downAt = bound_.vertical._MaxX._Max(worksheetPart, _mergedNoPre);


			return new _sheet._coord_._row.val.Range(
				new _sheet._coord_._row.Val(_mergedNoPre.RowIndex)

				,
				downAt
			);

		}
		static public _sheet._coord_._row.val.Range _Get(
			Worksheet worksheetPart
			,
			Row _mergedNoPre
		)
		{
			var downAt = bound_.vertical._MaxX._Max(worksheetPart, _mergedNoPre);


			return new _sheet._coord_._row.val.Range(
				new _sheet._coord_._row.Val(_mergedNoPre.RowIndex)

				
				,
				downAt
			);

		}

		public static _sheet._coord_._row.val.Range _Get(Worksheet worksheet, Val current)
		{
			var downAt = belt.range.bound_.vertical._MaxX._Max(worksheet, current);

			return new _sheet._coord_._row.val.Range(
				new _sheet._coord_._row.Val(current)

				
				,
				downAt
			);


		}

		

		static public _sheet._coord_._row.val.Range _Get(
			Worksheet worksheetPart
			,
			_sheet._coord_._row.Val _mergedNoPre
		)
		{
			var downAt = belt.range.bound_.vertical._MaxX._Max(worksheetPart, _mergedNoPre);


			return new _sheet._coord_._row.val.Range(
				_mergedNoPre

				
				,
				downAt
			);

		}

	}
}