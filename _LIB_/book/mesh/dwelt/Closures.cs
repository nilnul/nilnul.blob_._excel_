using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet.dwelt
{
	/// <summary>
	/// 
	/// </summary>
	public class Closures
	{

		private SheetData _sheetData;

		public SheetData sheetData
		{
			get { return _sheetData; }
			set { _sheetData = value; }
		}

		private nilnul.obj.matrix.block.Set _blocks;

		public nilnul.obj.matrix.block.Set blocks
		{
			get { return _blocks; }
			set { _blocks = value; }
		}




	}
}
