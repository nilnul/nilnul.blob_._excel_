using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.obj._matrix.coord.duo.be_.le.vow;
using nilnul.obj.matrix;

namespace nilnul.blob_.excel.doc.sheet
{
	/// <summary>
	/// the sheetData. a block
	/// </summary>
	public class Dwelt:
		nilnul.obj.matrix.block_.Diag
		,
		
		BlockI
	{
		private Sheet _sheet;

		public Sheet sheet
		{
			get { return _sheet; }
			set { _sheet = value; }
		}

		public Dwelt(Sheet sheet):base(  nilnul.fs.excel.doc.sheet. dwelt._AsDiagX.Get(sheet.worksheet))
		{
			_sheet = sheet;
		}

		public Dwelt( WorkbookPart workbookPart, Worksheet worksheet):this(new Sheet(workbookPart,worksheet))
		{
			
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
