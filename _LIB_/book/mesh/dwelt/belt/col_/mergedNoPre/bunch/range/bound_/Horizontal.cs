using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using nilnul.obj._matrix._coord_._col;
using nilnul.obj._matrix._coord_._row;
using nilnul.obj._mesh._coord.col_;
using nilnul.obj.matrix.block;
using nilnul.objs.be_;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.doc.sheet.dwelt.belt.col_.mergedNoPre.bunch.range.bound_
{
	/// <summary>
	/// get the outer colSpan of this col; this col has no cell merged with previous col.
	/// </summary>
	/// 
	static public class _HorizontalX
	{


		
		static public nilnul.blob_.excel.doc._sheet._coord.col.Bound _Get(
			Worksheet worksheetPart
			,
			fs.excel.doc._sheet._coord_._col.Val _mergedNoPre
		)
		{
			var rightAt = bound_._horizontal._MaxX._Max(worksheetPart, _mergedNoPre);


			return new nilnul.blob_.excel.doc._sheet._coord.col.Bound(
				_mergedNoPre

				,
				rightAt
			);

		}

		internal static nilnul.blob_.excel.doc._sheet._coord.col.Bound _Get(Worksheet worksheet, CaptialLetter current)
		{
			return _Get(worksheet , 
				new fs.excel.doc._sheet._coord_._col.Val(current.toNum().toBigint()+1)
				);
		}

		public static nilnul.blob_.excel.doc._sheet._coord.col.Bound _Get(
			Worksheet worksheet, obj._matrix._coord_._col.ValI current
		)
		{

			return _Get(worksheet, new fs.excel.doc._sheet._coord_._col.Val(current));

		}

		static public nilnul.blob_.excel.doc._sheet._coord.col.Bound _Get(
			WorksheetPart worksheetPart
			,
			fs.excel.doc._sheet._coord_._col.Val _mergedNoPre
		)
		{
			return _Get(worksheetPart.Worksheet,_mergedNoPre);

		}
		

		

		

	}
}