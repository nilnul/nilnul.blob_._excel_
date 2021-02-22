using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet
{
	/// <summary>
	/// non neg block. (block can be negative in that you drag in reversed way)
	/// </summary>
	/// 
	[Obsolete()]
	public class MergeCel

	{
		/// <summary>
		/// normalized to matrix.block and the sheet is regarded as matrix.
		/// </summary>
		/// <param name="rangeReference"></param>
		/// <returns></returns>
		static public nilnul.obj.matrix.Block Parse(string rangeReference)
		{
			var split = Regex.Matches(rangeReference,
				@"^
					(?<b>[^:]*)
					:
					(?<e>[^:]*)
				$"
				,
				RegexOptions.IgnorePatternWhitespace | RegexOptions.ExplicitCapture
			);

			var begin = split[0].Groups["b"].Value;
			var end = split[0].Groups["e"].Value;


			return new nilnul.obj.matrix.Block(
				sheet.dwelt.row.cel._key.Ret.Parse(begin)
				,
				dwelt.row.cel._key.Ret.Parse(end)

			);

			//throw new NotImplementedException();

		}



		static public nilnul.obj.matrix.Block Cast(DocumentFormat.OpenXml.Spreadsheet.MergeCell mergeCell)
		{
			return Parse(
			mergeCell.Reference

				);
		}


	}
}
