using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.doc.sheet.dwelt
{
	/// <summary>
	///a belt over closures is different from a belt over cells in that:
	///	1) a closure must not fall onto plural belts.
	///	2) a belt cannot be further divided into plural belts.
	/// </summary>
	public class Belt
	{
		private Dwelt _dwelt;

		public Dwelt dwelt
		{
			get { return _dwelt; }
			set { _dwelt = value; }
		}

		private nilnul.num.ord_.oneBased.Bound _rows;

		public nilnul.num.ord_.oneBased.Bound rows
		{
			get { return _rows; }
			set { _rows = value; }
		}

		

		public Belt(Dwelt dwelled, nilnul.num.ord_.oneBased.Bound rows)
		{
			_dwelt = dwelled;
			_rows = rows;
		}


		public Belt(WorkbookPart workbookPart, Worksheet worksheet, nilnul.num.ord_.oneBased.Bound rows)
			:this(new Dwelt(workbookPart,worksheet),rows)
		{
			
		}





		public nilnul.blob_.excel.doc._sheet._coord.col.Bound cols {
			get {
				return new blob_.excel.doc._sheet._coord.col.Bound( _dwelt.colRange );
			}
		}

		//public IEnumerable<nilnul.obj.matrix.block_.DiagI> partition() {

		//	//iterate each row
		//}
		


		


	}
}
