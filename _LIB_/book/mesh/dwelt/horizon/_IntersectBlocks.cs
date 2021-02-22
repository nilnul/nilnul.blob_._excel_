using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.obj.matrix;
using nilnul.obj._matrix._coord_._row;

namespace nilnul.blob_.excel.doc.sheet.dwelt.row
{
	static public class _IntersectConfinesX
	{
		static public nilnul.obj.matrix.block.Set GetSet(
			Worksheet worksheet
			,
			nilnul.fs.excel.doc._sheet._coord_._row.Val row
		)
		{

			return new nilnul.obj.matrix.block.Set(
				GetIntersectBlocks(worksheet,row)
			);

		}


		/// <summary>
		/// repeated elements
		/// </summary>
		/// <param name="worksheet"></param>
		/// <param name="row"></param>
		/// <returns></returns>
		public static IEnumerable<nilnul.obj.matrix.block_.Diag> GetIntersectBlocks(Worksheet worksheet, ValI row)
		{
			// get all the cells.
			var cels = nilnul.fs.excel.doc.sheet.dwelt.row._CelsX.Enumerate(worksheet, row);

			//get all the blocks
			var blocks = cels.Select(
				c => nilnul.fs.excel.doc.sheet.dwelt.cel.to_._ClosureX.ByMergeCellReference(worksheet, c)
			);
			return blocks;

			//throw new NotImplementedException();
		}

		static public nilnul.obj.matrix.block.Set GetSet(
			Worksheet worksheet
			,
			ValI row
		)
		{

			return new nilnul.obj.matrix.block.Set(
				GetIntersectBlocks(worksheet,row)
			);

		}


		[Obsolete(nameof(GetIntersectBlocks))]
		static public IEnumerable<nilnul.obj.matrix.block_.Diag> GetIntersectBlocks(
			Worksheet worksheet
			,
			nilnul.fs.excel.doc._sheet._coord_._row.Val row
		)
		{

			// get all the cells.
			var cels = nilnul.fs.excel.doc.sheet.dwelt.row._CelsX.Enumerate(worksheet, row);

			//get all the blocks
			var blocks = cels.Select(
				c => nilnul.fs.excel.doc.sheet.dwelt.cel.to_._ClosureX.ByMergeCellReference(worksheet, c)
			);
			return blocks;

		}


	}
}
