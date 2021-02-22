using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.obj._matrix.coord.duo.be_.le.vow;

namespace nilnul.fs.excel.doc.sheet
{
	/// <summary>
	/// the sheetData. a block
	/// </summary>
	public class Dwelt:
		nilnul.obj.matrix.block_.Diag
		,
		
		BlockI
	{
		private Worksheet _worksheet;

		public Worksheet worksheet
		{
			get { return _worksheet; }
			set { _worksheet = value; }
		}

		/// <summary>
		/// 
		/// </summary>
		/// <remarks>
		/// need this to get shared string table
		/// </remarks>
		private WorkbookPart _workbookPart;

		public WorkbookPart workbookPart
		{
			get { return _workbookPart; }
			set { _workbookPart = value; }
		}

		public Dwelt( WorkbookPart workbookPart, Worksheet worksheet):base(  dwelt._AsDiagX.Get(worksheet))
		{
			_workbookPart = workbookPart;
			_worksheet = worksheet;
		}

		
		

		

		static public Dwelt _Create(SpreadsheetDocument doc, Worksheet _worksheet__mustBeInDoc) {
			return new Dwelt(
				doc.WorkbookPart
				,
				_worksheet__mustBeInDoc
			);
		}



	}
}
