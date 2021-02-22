using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using N = nilnul.NumI;

namespace nilnul.fs.excel.doc.sheet.dwelt.row.span_
{
	/// <summary>
	/// a collection of adjacent rows; at least one row must be present.
	/// </summary>
	/// <remarks>
	/// if a merge cell spans a collection of rows, these rows shall be placed in the same span
	/// </remarks>
	public class BasicExtra
		: num.range_.BasicExtra
	{
		public BasicExtra(Num basic, Num extra):base(basic, extra)
		{

		}
	}
}
