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

namespace nilnul.fs.excel.doc._sheet
{
	public class Coord :
		nilnul.obj._matrix.Coord<  _coord_._col.Val, _coord_._row.Val>
		//,
		//_coord_.RowI
		//,
		//_coord_.ColI




	{

		public Coord(_coord_._col.Val col, _coord_._row.Val row) : base(row, col)
		{
		}

		public Coord( _coord_._row.Val row, _coord_._col.Val col) : base(row, col)
		{
		}

		public Coord(CoordI_colValRowVal x):this(
			new _coord_._col.Val((x as CoordI).col)
			,
			new _coord_._row.Val((x as CoordI).row)
		)
		{
		}

		public Coord(CoordI begin):this(begin.row,begin.col)
		{
		}

		

		public Coord(obj._matrix._coord_._row.ValI row, obj._matrix._coord_._col.ValI col)
			:base(
				 new _coord_._col.Val(col)
				 
				,
				 
				 new _coord_._row.Val(row))
		{
		}

	

		public Coord(obj._matrix._coord_._col.ValI col, obj._matrix._coord_._row.ValI row):this(row,col)
		{
		}

		public override string ToString()
		{
			return $"{col}{row}";
		}

		static public Coord Parse(string s)
		{
			//versioned.

			var versioned = nilnul.txt_.Versioned.Parse(s);

			return new Coord(
				 new doc._sheet._coord_._row.Val((int)versioned.version.val)
				,
				_coord_._col.Val.Parse_bigEndian(versioned.prefix)

			);
		}

		static public Coord CreateFroCell(Cell cell)
		{
			return Parse(cell.CellReference);
		}
	}
}
