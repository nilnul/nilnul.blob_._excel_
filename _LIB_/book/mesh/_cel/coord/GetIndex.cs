using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet.row.cell
{
	public class GetIndex
	{
		public static int GetRowIndexFromCellAddress(string cellAddress)
		{
			//  Convert an Excel CellReference column into a 1-based row index
			//  eg "D42"  ->  42
			//     "F123" ->  123
			string rowNumber = System.Text.RegularExpressions.Regex.Replace(cellAddress, "[^0-9 _]", "");
			return int.Parse(rowNumber);
		}



	}
}
