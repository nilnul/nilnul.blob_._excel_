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

namespace nilnul.blob_.excel.doc.sheet.dwelt.belt.col_.mergedNoPre.bunch.range.bound_._horizontal._max
{
	/// <summary>
	/// get the outer rowSpan of this row; this row has no cell merged with previous row.
	/// </summary>
	/// 
	static public class _RightAtX
	{
		

		static public fs.excel.doc._sheet._coord_._col.Val __RightAt(
			Worksheet worksheet
			,
			fs.excel.doc._sheet._coord_._col.Val col
		)
		{
			// get all the cells.
			var reifiedCells = sheet.dwelt.col.cels_._ReifiedX.Ordered(worksheet,col);

			//get all the blocks
			var closures = reifiedCells.Select(c => nilnul.fs.excel.doc.sheet.dwelt.cel.to_._ClosureX.ByMergeCellReference(worksheet, c));
			//
			var maxCol = closures.Select(b => b.end.col).Aggregate(
				new nilnul.fs.excel.doc._sheet._coord_._col.Val(	//if there are no cells.
					col
				) as nilnul.obj._matrix._coord_._col.ValI
				,
				(a, c) => nilnul.obj._matrix._coord.col.op_.binary_.Max.Singleton.op(
					a, c
				)
			);
			return new nilnul.fs.excel.doc._sheet._coord_._col.Val( maxCol);
		}

		

		//static public nilnul.fs.excel.doc._sheet._coord_._col.Val __DownAt(
		//	WorksheetPart worksheetPart
		//	,
		//	Row row
		//)
		//{

		//	return __DownAt(worksheetPart.Worksheet, row);
		//	// get all the cells.
		


		//}

		//static public nilnul.fs.excel.doc._sheet._coord_._col.Val __DownAt(
		//	Worksheet worksheet
		//	,
		//	Col row
		//)
		//{
		//	return __DownAt(worksheet, new nilnul.fs.excel.doc._sheet._coord_._col.Val(row));
		//	//// get all the cells.
		//	//var reifiedCells = sheet.dwelt.row.cels_._ReifiedX.Enumerate(row);


		//	////get all the blocks
		//	//var closures = reifiedCells.Select(c => sheet.dwelt.cel.to_._ClosureX.ByMergeCellReference(worksheet, c));

		//	////
		//	//var maxRow = closures.Select(b => b.end.row).Aggregate(
		//	//	new nilnul.fs.excel.doc._sheet._coord_._row.Val(	//if there are no cells.
		//	//		row
		//	//	) as nilnul.obj._matrix._coord_._row.ValI

		//	//	,
		//	//	(a, c) => nilnul.obj._matrix._coord_._row.val.combine_.Max.Singleton.combine(
		//	//		a, c
		//	//	)
		//	//);

		//	//return new _sheet._coord_._row.Val( maxRow);
		//}

			/// <summary>
			/// may have repeated.
			/// </summary>
			/// <param name="worksheet"></param>
			/// <param name="row"></param>
			/// <returns></returns>
		static public IEnumerable< nilnul.obj.matrix.BlockI> __RightAt_retClosures(
			Worksheet worksheet
			,
			nilnul.fs.excel.doc._sheet._coord_._col.Val row
		)
		{
			// get all the cells.
			var reifiedCells = sheet.dwelt.col.cels_._ReifiedX.Ordered(worksheet,row);
			//get all the blocks
			var closures = reifiedCells.Select(
				c => nilnul.fs.excel.doc.sheet.dwelt.cel.to_._ClosureX.ByMergeCellReference(worksheet, c)
			);

			return closures;

		}






	}
}