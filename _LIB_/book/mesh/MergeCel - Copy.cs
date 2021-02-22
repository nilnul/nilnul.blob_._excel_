using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet.dwelt.closure_
{
	/// <summary>
	/// </summary>
	static public class _MergeCelX

	{
		/// <summary>
		/// </summary>
		/// <param name="rangeReference"></param>
		/// <returns></returns>
		static public nilnul.obj.matrix.block_.Diag Parse(string rangeReference)
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


			return new nilnul.obj.matrix.block_.Diag(


				_sheet.Coord.Parse(begin)
				,
				_sheet.Coord.Parse(end)

			);

			//throw new NotImplementedException();

		}



		static public nilnul.obj.matrix.block_.Diag ToDiag(DocumentFormat.OpenXml.Spreadsheet.MergeCell mergeCell)
		{
			return Parse(
				mergeCell.Reference

			);
		}


	}
}
