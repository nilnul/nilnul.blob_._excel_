using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet.dwelt.closures
{
	/// <summary>
	///a belt over closures is different from a belt over cells in that:
	///	1) a closure must not fall onto plural belts.
	///	2) a belt cannot be further divided into plural belts.
	/// </summary>
	public class Belt
	{
		private Dwelt _dwelled;

		public Dwelt dwelled
		{
			get { return _dwelled; }
			set { _dwelled = value; }
		}

		private nilnul.obj._matrix._coord_._row.val.Range _rows;

		public nilnul.obj._matrix._coord_._row.val.Range rows
		{
			get { return _rows; }
			set { _rows = value; }
		}

		

		public Belt(Dwelt dwelled, nilnul.obj._matrix._coord_._row.val.Range range)
		{
			_dwelled = dwelled;
			_rows = range;
		}

		public nilnul.obj._matrix._coord_._col.val.Range cols {
			get {
				return _dwelled.colRange;
			}
		}

		//public IEnumerable<nilnul.obj.matrix.block_.DiagI> partition() {

		//	//iterate each row
		//}
		


		


	}
}
