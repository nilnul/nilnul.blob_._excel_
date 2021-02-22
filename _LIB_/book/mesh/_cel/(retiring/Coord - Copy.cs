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

namespace nilnul.blob_.excel.doc._sheet
{
	[Obsolete()]
	public class Coord :
		nilnul.obj._mesh.Coord<nilnul.obj._mesh._coord.row_.Positive, nilnul.obj._mesh._coord.col_.CaptialLetter>
	//,
	//_coord_.RowI
	//,
	//_coord_.ColI




	{

		public nilnul.obj._matrix.CoordI toCoord()
		{
			return new nilnul.obj._matrix.Coord(
				new nilnul.obj._matrix._coord_._row.Val(this.row.toNum())
				,
				new nilnul.obj._matrix._coord_._col.Val(this.col.toNum())

			);
		}


		public Coord(CaptialLetter col, obj._mesh._coord.row_.Positive row) : base(col, row)
		{
		}

		public Coord(obj._mesh._coord.row_.Positive row, CaptialLetter col) : base(row, col)
		{
		}

		public Coord(obj._mesh.Coord x) : this(x.row, x.col)
		{
		}

		public Coord(obj._mesh._coord.RowI row1, obj._mesh._coord.ColI col1)
			: this(
				new nilnul.obj._mesh._coord.row_.Positive(row1)
				 ,
				new obj._mesh._coord.col_.CaptialLetter(col1)
			)
		{
		}



		public Coord(CoordI_colValRowVal x) : this(
			new CaptialLetter(
				new nilnul.num.Ord1(
					((ColI<nilnul.obj._matrix._coord_._col.ValI>)x).col.toNum()
				)
			)
			,
			 new nilnul.obj._mesh._coord.row_.Positive(
				 nilnul.num.ord_.OneBased.FroNilBased(

					new nilnul.Num(
							((RowI<nilnul.obj._matrix._coord_._row.ValI>)x).row.toNum()
							)
				)

				)
		)
		{
		}

		public override string ToString()
		{
			return $"{col}{row}";
		}
		[Obsolete(nameof(blob_.excel.doc.sheet.cel._CoordX))]
		static public Coord Parse(string s)
		{
			//versioned.

			var versioned = nilnul.txt_.Versioned.Parse(s);

			return new Coord(
				  obj._mesh._coord.row_.Positive.FroOneBased(new num_.Positive1(versioned.version.val))
				,
				CaptialLetter.Parse(versioned.prefix)

			);
		}

			[Obsolete(nameof(blob_.excel.doc.sheet.cel._CoordX))]
	static public Coord CreateFroCell(Cell cell)
		{
			return Parse(cell.CellReference);
		}
	}
}
