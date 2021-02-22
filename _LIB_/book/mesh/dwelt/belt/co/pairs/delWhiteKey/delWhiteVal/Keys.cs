using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using nilnul.fs.excel.doc._sheet._coord_._row;
using DocumentFormat.OpenXml.Spreadsheet;

namespace nilnul.blob_.excel.doc.sheet.dwelt.belt.co.pairs.delWhiteKey.delWhiteVal
{
	/// <summary>
	/// </summary>
	static public class _KeysX
	{
		

		public static IEnumerable<string> Keys(
			SpreadsheetDocument doc
			,
			Worksheet worksheet
			,
			nilnul.obj._matrix._coord_._row.ValI row
			,
			nilnul.obj._matrix._coord_._row.ValI rowFilter
		)
		{
			return 
					pairs.delWhiteKey._DelWhiteValX.ExceptWhiteVal(doc, worksheet, row,rowFilter).Select(x=>x.Item1?.Trim())
				;
		}



	}
}