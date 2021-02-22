using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet.dwelt.belt.block_.mergeCel.cel_
{
	public class LeftUpper
	{
		/// <summary>
		/// get the coord of the upperLeft cel
		/// </summary>
		/// <param name="mergeCell"></param>
		/// <returns></returns>
		static public _sheet.Coord GetCoord(MergeCell mergeCell) {

			return new _sheet.Coord( block_.MergeCelX.ToDiag(mergeCell).begin );

		}

		
	}
}
