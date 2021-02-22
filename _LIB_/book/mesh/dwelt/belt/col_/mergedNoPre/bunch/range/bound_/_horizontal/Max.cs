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
using nilnul.fs.excel;

namespace nilnul.blob_.excel.doc.sheet.dwelt.belt.col_.mergedNoPre.bunch.range.bound_._horizontal
{
	/// <summary>
	/// get the outer rowSpan of this row; this row has no cell merged with previous row.
	/// </summary>
	/// 
	static public class _MaxX
	{
		static public fs.excel.doc._sheet._coord_._col.Val _Max(
			Worksheet worksheet
			,
			fs.excel.doc._sheet._coord_._col.Val _mergedNoPre

		)
		{

			var rightAt = _max._RightAtX.__RightAt(worksheet, _mergedNoPre);


			if ( nilnul.obj._matrix._coord_._col.val.Eq.Singleton.Equals(
				rightAt,  
				new fs.excel.doc._sheet._coord_._col.Val(
					_mergedNoPre
				)
				)
				
			)
			{
				return rightAt;
			}

			return _Max(worksheet, rightAt);

		}

		//static public fs.excel.doc._sheet._coord_._row.Val _Max(
		//	Worksheet worksheet
		//	,
		//	Row _mergedNoPre

		//)
		//{
		//	return _Max(worksheet, new fs.excel.doc._sheet._coord_._row.Val(_mergedNoPre));

			

		//}

		//static public fs.excel.doc._sheet._coord_._col.Val _Max(
		//	WorksheetPart worksheetPart
		//	,
		//	Row _mergedNoPre
		//)
		//{
		//	return _Max(worksheetPart.Worksheet, _mergedNoPre);
		//}

		static public fs.excel.doc._sheet._coord_._col.Val _Max(
			Worksheet worksheet
			,
			nilnul.obj._matrix._coord_._col.Val _mergedNoPre

		)
		{
			return _Max(worksheet,  new nilnul.fs.excel.doc._sheet._coord_._col.Val(_mergedNoPre));

		}







		public static nilnul.fs.excel.doc._sheet._coord_._col.Val _Max(WorksheetPart worksheetPart, nilnul.fs.excel.doc._sheet._coord_._col.Val rowIndex)
		{

			return _Max(worksheetPart.Worksheet, rowIndex);


		}

		





	}
}