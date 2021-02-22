using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using nilnul.fs.excel.doc._sheet._coord_._row;
using DocumentFormat.OpenXml.Spreadsheet;

namespace nilnul.blob_.excel.doc.sheet.dwelt.belt.co.pairs
{
	/// <summary>
	/// </summary>
	static public class _DelWhiteKeyX
	{
		
		/// <summary>
		/// white keys are ignored, together with their values (if there is no key but value, it's ignored)
		/// </summary>
		/// <param name="doc"></param>
		/// <param name="worksheet"></param>
		/// <param name="row"></param>
		/// <returns></returns>
		public static IEnumerable<nilnul.txt.Duo> ExceptWhiteKey(
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
					_PairsX.GetTxtDuoS(doc, worksheet, row,rowFilter).Where(
						x => nilnul.txt.be_.NonWhite.Singleton.be(x.Item1)  //white keys are ignored
					)
				;
		}

	
	}
}