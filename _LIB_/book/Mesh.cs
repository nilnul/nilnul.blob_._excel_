using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.obj._matrix.coord.duo.be_.le.vow;
using nilnul.obj.matrix;

namespace nilnul.blob_.excel.book
{
	/// <summary>
	/// </summary>
	public class Mesh
		
	{
		private WorksheetPart _worksheetPart;

		public WorksheetPart worksheetPart
		{
			get { return _worksheetPart; }
			//set { _worksheet = value; }
		}


		/// <summary>
		/// 
		/// </summary>
		/// <remarks>
		/// need this to get shared string table
		/// </remarks>
		private SpreadsheetDocument _book;
		private SpreadsheetDocument workbookPart1;
		private Worksheet worksheets;

		public SpreadsheetDocument book {
			get {
				return _book;
			}
		}

		public WorkbookPart workbookPart
		{
			get { return _book.WorkbookPart; }
			//set { _workbookPart = value; }
		}

		
		public Mesh( SpreadsheetDocument doc, WorksheetPart _worksheet__mustBeInDoc) {
			_book = doc;
			_worksheetPart = _worksheet__mustBeInDoc;
		}

		public Mesh(SpreadsheetDocument workbookPart1, Worksheet worksheets)
		{
			this.workbookPart1 = workbookPart1;
			this.worksheets = worksheets;
		}
	}
}
