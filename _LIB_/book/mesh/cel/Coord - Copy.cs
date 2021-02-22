using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.obj._matrix;
using nilnul.fs.excel.doc._sheet._coord_;
using DocumentFormat.OpenXml.Spreadsheet;
using nilnul.obj._matrix._coord_._col;
using nilnul.obj._matrix._coord_._row;
using nilnul.num.ord_;
using nilnul.num.ord_.positive;
using nilnul.obj._mesh._coord.col_;
using nilnul.obj._mesh._coord.row_;
using nilnul.obj._mesh;
using nilnul.obj._mesh._coord;
using nilnul.obj._matrix._coord_;

namespace nilnul.blob_.excel.doc.sheet.cel
{
	static public class _CoordX
	{

		static public sheet._cel.Coord Parse(string s)
		{
			//versioned.
			var versioned = nilnul.txt_.Versioned.Parse(s);

			return new sheet._cel.Coord(
				nilnul.num.ord_.oneBased_.bijective_.UpperLetter.Parse(
					versioned.prefix
				)
				,
				nilnul.num.ord_.OneBased.FroOneBased(new num_.Positive1(versioned.version.val))
			);
		}

		static public sheet._cel.Coord CreateFroCell(Cell cell)
		{
			return Parse(cell.CellReference);
		}
	}
}
