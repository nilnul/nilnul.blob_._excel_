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

namespace nilnul.fs.excel.doc.sheet.dwelt.row_.mergedNoPre
{
	/// <summary>
	/// get the outer rowSpan of this row; this row has no cell merged with previous row.
	/// </summary>
	/// 

		[Obsolete()]
	public class Belt_retRange
	{



		static public _sheet._coord_._row.val.Range _Get(
			WorksheetPart worksheetPart
			,
			Row _mergedNoPre
		)
		{
			var downAt = _DowAtRecur(worksheetPart, _mergedNoPre);


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
			var downAt = _DowAtRecur(worksheetPart, _mergedNoPre);


			return new _sheet._coord_._row.val.Range(
				new _sheet._coord_._row.Val(_mergedNoPre.RowIndex)

				
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
			var downAt = _DowAtRecur(worksheetPart, _mergedNoPre);


			return new _sheet._coord_._row.val.Range(
				_mergedNoPre

				
				,
				downAt
			);

		}

	


		static public _sheet._coord_._row.Val _DowAtRecur(
			WorksheetPart worksheetPart
			,
			Row _mergedNoPre

		)
		{

			return _DowAtRecur(worksheetPart.Worksheet, _mergedNoPre);

			

		}


		static public _sheet._coord_._row.Val _DowAtRecur(
			Worksheet worksheet
			,
			Row _mergedNoPre

		)
		{

			var downAt = _DownAt(worksheet, _mergedNoPre);


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

			return _DowAtRecur(worksheet, downAt);

		}
	



		static public _sheet._coord_._row.Val _DowAtRecur(
			Worksheet worksheet
			,
			_sheet._coord_._row.Val _mergedNoPre

		)
		{

			var downAt = _DownAt(worksheet, _mergedNoPre);


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

			return _DowAtRecur(worksheet, downAt);

		}



		public static _sheet._coord_._row.Val _DowAtRecur(WorksheetPart worksheetPart, _sheet._coord_._row.Val rowIndex)
		{

			return _DowAtRecur(worksheetPart.Worksheet, rowIndex);


		}

		

		static public _sheet._coord_._row.Val _DownAt(
			WorksheetPart worksheetPart
			,
			Row row
		)
		{

			return _DownAt(worksheetPart.Worksheet, row);
			// get all the cells.
		


		}

		static public _sheet._coord_._row.Val _DownAt(
			Worksheet worksheet
			,
			Row row
		)
		{

			// get all the cells.
			var reifiedCells = sheet.dwelt.row.cels_._ReifiedX.Enumerate(row);


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

		static public _sheet._coord_._row.Val _DownAt(
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


		static public IEnumerable< nilnul.obj.matrix.BlockI> _DownAt_retClosures(
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