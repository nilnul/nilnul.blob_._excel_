using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.doc.sheet.dwelt.cel
{
	/// <summary>
	/// x, col
	/// </summary>
	static public class _HorizontalIndexX
	{

		/// <summary>
		///  Given a cell name, parses the specified cell to get the column name.
		/// </summary>
		/// <param name="cellName"></param>
		/// <returns></returns>
		static  public string ColName(string cellName)
		{
			// Create a regular expression to match the column name portion of the cell name.
			Regex regex = new Regex("[A-Za-z]+");
			Match match = regex.Match(cellName);

			return match.Value;
		}


	}
}
