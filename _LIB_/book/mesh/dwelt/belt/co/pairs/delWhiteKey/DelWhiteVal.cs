using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using nilnul.fs.excel.doc._sheet._coord_._row;
using DocumentFormat.OpenXml.Spreadsheet;

namespace nilnul.blob_.excel.doc.sheet.dwelt.belt.co.pairs.delWhiteKey
{
	/// <summary>
	/// </summary>
	static public class _DelWhiteValX
	{
		

		public static IEnumerable<nilnul.txt.Duo> ExceptWhiteVal(
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
					_DelWhiteKeyX.ExceptWhiteKey(doc, worksheet, row,rowFilter).Where(
						x => nilnul.txt.be_.NonWhite.Singleton.be(x.Item2)  //white keys are ignored
					)
				;
		}

		
	}
}