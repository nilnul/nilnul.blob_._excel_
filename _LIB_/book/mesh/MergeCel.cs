using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.book.mesh
{
	/// <summary>
	/// </summary>
	static public class _CelMergedX

	{
		/// <summary>
		/// </summary>
		/// <param name="rangeReference"></param>
		/// <returns></returns>
		static public nilnul.obj.mesh.Bloc Parse(string rangeReference)
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


			return new nilnul.obj.mesh.Bloc(


				_cel.Coord.Parse(begin)
				,
				_cel.Coord.Parse(end)

			);


			//throw new NotImplementedException();

		}



		static public nilnul.obj.mesh.Bloc ToDiag(DocumentFormat.OpenXml.Spreadsheet.MergeCell mergeCell)
		{
			return Parse(
				mergeCell.Reference

			);
		}


	}
}
