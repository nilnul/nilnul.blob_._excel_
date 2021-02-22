using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using nilnul.fs.excel.doc.sheet.dwelt.closure_;
using nilnul.fs.excel.doc.sheet.dwelt.closures_;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.obj.matrix.block_;
using nilnul.obj._mesh;

namespace nilnul.blob_.excel.doc.sheet.dwelt.cel
{
	/// <summary>
	/// the cell itself or the bounding mergeCel
	/// </summary>
	static public class _ConfineX
	{
		static public nilnul.obj.matrix.block_.Diag ByMergeCellReference(WorksheetPart wsPart, 
			_sheet.Coord cell_ofSameSheet)
		{
			return ByMergeCellReference(wsPart.Worksheet,cell_ofSameSheet);
		}

		internal static Diag ByMergeCellReference(Dwelt dwelt, Coord c)
		{

			return ByMergeCellReference(
				dwelt.sheet.worksheet
				,
				new _sheet.Coord(c)
			);
		}

		static public nilnul.obj.matrix.block_.Diag ByMergeCellReference(
			Worksheet worksheet
			, 
			_sheet.Coord cel
		)
		{
			var mergeCells = nilnul.fs.excel.doc. sheet.dwelt.closures_._MergeCels.Enumerate(worksheet).Where(
				x => _MergeCelX.ToDiag(x).contain(
					cel.toCoord()
				)
			);

			if (mergeCells.Any())
			{
				return _MergeCelX.ToDiag(
					mergeCells.First()
				);
			}
			else
			{
				return nilnul.obj.matrix.block_.Diag.CreateFroCel(cel.toCoord());
			}
		}


		static public nilnul.obj.matrix.block_.Diag ByMergeCellReference(
			Worksheet worksheet
			, 
			nilnul.obj._matrix.CoordI cel
		)
		{
			var mergeCells = nilnul.fs.excel.doc. sheet.dwelt.closures_._MergeCels.Enumerate(worksheet).Where(
				x => _MergeCelX.ToDiag(x).contain(
					cel
				)
			);

			if (mergeCells.Any())
			{
				return _MergeCelX.ToDiag(
					mergeCells.First()
				);
			}
			else
			{
				return nilnul.obj.matrix.block_.Diag.CreateFroCel(cel);
			}
		}



		public  static Diag ByMergeCellReference(Worksheet worksheet, Cell c)
		{
			return ByMergeCellReference(worksheet,  _sheet.Coord.CreateFroCell(c));
			throw new NotImplementedException();
		}
	}
}
