using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.doc.sheet.dwelt.cel
{
	static public class _VertialIndexX
	{

		// Given a cell name, parses the specified cell to get the row index.
		/// <summary>
		/// starting from 1.
		/// </summary>
		/// <param name="cellName"></param>
		/// <returns></returns>
		public static uint RowIndex(string cellName)
		{
			// Create a regular expression to match the row index portion the cell name.
			Regex regex = new Regex(@"\d+");
			Match match = regex.Match(cellName);

			return uint.Parse(match.Value);
		}
	}
}
