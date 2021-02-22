using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.num;
using DocumentFormat.OpenXml.Spreadsheet;
using nilnul.fs.excel.doc.sheet.dwelt.row.cel;
using nilnul.fs.excel.doc.sheet.dwelt.row.cel._key;
using nilnul.num.duo;
using nilnul.fs.excel.doc.sheet.dwelt.row.cel._key._ret;

namespace nilnul.fs.excel.doc.sheet.dwelt.row.cel._key
{
	/// <summary>
	/// the value of a key
	/// A1， AH36, etc
	/// </summary>
	/// 
	[Obsolete(nameof(excel.doc._sheet.Coord))]
	public class Ret : 

		nilnul.obj._matrix_.cell.RowMajorCoordI,
		
		_key_.ColI
	{

		private Col _col;

		public Col col
		{
			get { return _col; }
			set { _col = value; }
		}

		private _ret.Row _row;

		public _ret.Row row
		{
			get { return _row; }
			set { _row = value; }
		}

		public Coord coords
		{
			get
			{
				return new Coord(
					_col.index
					,
					_row.index


				);
				//throw new NotImplementedException();
			}
		}

		

		public Ret(Col col, _ret.Row row)
		{
			_col = col;
			_row = row;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <returns></returns>
		public override string ToString()
		{
			return $"{_col}{_row}";
		}

		static public Ret Parse(string s) {
			//versioned.

			var versioned = nilnul.txt_.Versioned.Parse(s);

			return new Ret(
				Col.Parse_bigEndian(versioned.prefix)
				,
				_ret.Row.CreateFroRow(versioned.version.val)
			);
		}

		static public Ret FroCell(Cell cell) {
			return Parse( cell.CellReference);
		}
	}
}
