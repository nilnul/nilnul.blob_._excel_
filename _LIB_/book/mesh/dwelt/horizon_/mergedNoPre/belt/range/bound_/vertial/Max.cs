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

namespace nilnul.fs.excel.doc.sheet.dwelt.row_.mergedNoPre.belt.range.bound_.vertical
{
	/// <summary>
	/// get the outer rowSpan of this row; this row has no cell merged with previous row.
	/// </summary>
	/// 
	static public class _MaxX
	{
		static public _sheet._coord_._row.Val _Max(
			Worksheet worksheet
			,
			_sheet._coord_._row.Val _mergedNoPre

		)
		{

			var downAt = __DownAt(worksheet, _mergedNoPre);


			if ( nilnul.obj._matrix._coord_._row.val.Eq.Singleton.Equals(
				downAt,  
				new _sheet._coord_._row.Val(
					_mergedNoPre
				)
				)
				
			)
			{
				return downAt;
			}

			return _Max(worksheet, downAt);

		}

		static public _sheet._coord_._row.Val _Max(
			Worksheet worksheet
			,
			Row _mergedNoPre

		)
		{
			return _Max(worksheet, new _sheet._coord_._row.Val(_mergedNoPre));

			//var downAt = _DownAt(worksheet, _mergedNoPre);

			//if ( nilnul.obj._matrix._coord_._row.val.Eq.Singleton.Equals(
			//	downAt,  
			//	new _sheet._coord_._row.Val(
			//		_mergedNoPre
			//	)
			//	)
				
			//)
			//{
			//	return downAt;
			//}

			//return _Max(worksheet, downAt);

		}

		static public _sheet._coord_._row.Val _Max(
			WorksheetPart worksheetPart
			,
			Row _mergedNoPre
		)
		{
			return _Max(worksheetPart.Worksheet, _mergedNoPre);
		}

		static public _sheet._coord_._row.Val _Max(
			Worksheet worksheet
			,
			nilnul.obj._matrix._coord_._row.Val _mergedNoPre

		)
		{
			return _Max(worksheet,  new nilnul.fs.excel.doc._sheet._coord_._row.Val(_mergedNoPre));

		}







		public static _sheet._coord_._row.Val _Max(WorksheetPart worksheetPart, _sheet._coord_._row.Val rowIndex)
		{

			return _Max(worksheetPart.Worksheet, rowIndex);


		}

		static public _sheet._coord_._row.Val __DownAt(
			Worksheet worksheet
			,
			_sheet._coord_._row.Val row
		)
		{
			// get all the cells.
			var reifiedCells = sheet.dwelt.row.cels_._ReifiedX.Enumerate(worksheet,row);
			//get all the blocks
			var closures = reifiedCells.Select(c => sheet.dwelt.cel.to_._ClosureX.ByMergeCellReference(worksheet, c));
			//
			var maxRow = closures.Select(b => b.end.row).Aggregate(
				new nilnul.fs.excel.doc._sheet._coord_._row.Val(	//if there are no cells.
					row
				) as nilnul.obj._matrix._coord_._row.ValI
				,
				(a, c) => nilnul.obj._matrix._coord_._row.val.combine_.Max.Singleton.combine(
					a, c
				)
			);
			return new _sheet._coord_._row.Val( maxRow);
		}

		

		static public _sheet._coord_._row.Val __DownAt(
			WorksheetPart worksheetPart
			,
			Row row
		)
		{

			return __DownAt(worksheetPart.Worksheet, row);
			// get all the cells.
		


		}

		static public _sheet._coord_._row.Val __DownAt(
			Worksheet worksheet
			,
			Row row
		)
		{
			return __DownAt(worksheet, new _sheet._coord_._row.Val(row));
			//// get all the cells.
			//var reifiedCells = sheet.dwelt.row.cels_._ReifiedX.Enumerate(row);


			////get all the blocks
			//var closures = reifiedCells.Select(c => sheet.dwelt.cel.to_._ClosureX.ByMergeCellReference(worksheet, c));

			////
			//var maxRow = closures.Select(b => b.end.row).Aggregate(
			//	new nilnul.fs.excel.doc._sheet._coord_._row.Val(	//if there are no cells.
			//		row
			//	) as nilnul.obj._matrix._coord_._row.ValI

			//	,
			//	(a, c) => nilnul.obj._matrix._coord_._row.val.combine_.Max.Singleton.combine(
			//		a, c
			//	)
			//);

			//return new _sheet._coord_._row.Val( maxRow);
		}


		static public IEnumerable< nilnul.obj.matrix.BlockI> __DownAt_retClosures(
			Worksheet worksheet
			,
			_sheet._coord_._row.Val row
		)
		{
			// get all the cells.
			var reifiedCells = sheet.dwelt.row.cels_._ReifiedX.Enumerate(worksheet,row);
			//get all the blocks
			var closures = reifiedCells.Select(
				c => sheet.dwelt.cel.to_._ClosureX.ByMergeCellReference(worksheet, c)
			);

			return closures;

		}






	}
}