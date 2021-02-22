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
using nilnul.num.ord_.oneBased_.bijective_;

namespace nilnul.blob_.excel.doc.sheet._cel
{
	public class Coord :
		nilnul.obj.mesh._cel.Coord<nilnul.num.ord_.oneBased_.bijective_.UpperLetter, nilnul.num.ord_.OneBased>
		//,
		//_coord_.RowI
		//,
		//_coord_.ColI



	{

		//public nilnul.obj._matrix.CoordI toCoord() {
		//	return new nilnul.obj._matrix.Coord(
		//		new nilnul.obj._matrix._coord_._row.Val(  this.row.toNum())
		//		,
		//		new nilnul.obj._matrix._coord_._col.Val(  this.col.toNum())
				
		//	);
		//}


		public Coord(CaptialLetter col, obj._mesh._coord.row_. Positive row) : base(new UpperLetter(col), row)
		{
		}

		static public Coord OvColRowNilBased(int col, int row)
		{
			return new Coord(
				 new obj.mesh._cel.Coord(col,row)
			);

		}

		//public Coord(obj._mesh._coord.row_. Positive row, CaptialLetter col) : this(new OneBased(row),new UpperLetter( col))
		//{
		//}

		//public Coord(obj._mesh.Coord x):this( x.col, x.row)
		//{
		//}

		//public Coord(obj._mesh._coord.RowI row1, obj._mesh._coord.ColI col1)
		//	:this(
		//		new nilnul.obj._mesh._coord.row_.Positive(row1 )
		//		 ,
		//		new obj._mesh._coord.col_.CaptialLetter(col1)
		//	)
		//{
		//}

		public Coord(UpperLetter col, OneBased row) : base(col, row)
		{
		}

		//public Coord(OneBased row, UpperLetter col) : base(row, col)
		//{
		//}

		public Coord(obj.mesh._cel.Coord x)
			:this(
				 new UpperLetter(x.col)
				 ,
				 new OneBased(x.row)
			 )
		{
			
		}

		//public Coord(CoordI_colValRowVal x):this(
		//	((RowI < nilnul.obj._matrix._coord_._row.ValI > )x).row
		//	, ((ColI<nilnul.obj._matrix._coord_._col.ValI>)x).col
		//)
		//{
		//}

		public override string ToString()
		{
			return $"{col}{row}";
		}

		static public Coord Parse(string s)
		{
			//versioned.

			var versioned = nilnul.txt_.Versioned.Parse(s);

			return new Coord(
				UpperLetter.Parse(versioned.prefix)
				,
				  obj._mesh._coord.row_.Positive.FroOneBased(new num_.Positive1( versioned.version.val ))
				

			);
		}

		static public Coord CreateFroCell(Cell cell)
		{
			return Parse(cell.CellReference);
		}
	}
}
