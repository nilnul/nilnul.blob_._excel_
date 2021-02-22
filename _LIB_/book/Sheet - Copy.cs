using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.obj._matrix.coord.duo.be_.le.vow;
using nilnul.obj.matrix;

namespace nilnul.blob_.excel.doc
{
	/// <summary>
	/// </summary>
	public class Sheet
		
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

		public Sheet( WorkbookPart workbookPart, Worksheet worksheet)
		{
			_workbookPart = workbookPart;
			_worksheet = worksheet;
		}

		public Sheet( WorkbookPart workbookPart, WorksheetPart worksheet):this(workbookPart,worksheet.Worksheet)
		{
			
		}


		static public Sheet _Create(SpreadsheetDocument doc, Worksheet _worksheet__mustBeInDoc) {
			return new Sheet(
				doc.WorkbookPart
				,
				_worksheet__mustBeInDoc
			);
		}



	}
}
