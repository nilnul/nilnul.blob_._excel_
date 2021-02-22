using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.obj._matrix._coord_._row;

namespace nilnul.fs.excel.doc._sheet._coord_._row.val
{
	public class Range:nilnul.obj._matrix._coord_._row.val.Range
	{
		private Range(ValI row1, ValI row2) : base(row1, row2)
		{
		}

		public Range( 
			Val a, Val b
		):base(a,b)
		{

		}

		static public Range CreateSingleton(Val a) {
			return new Range(a,a);
		}

		
	}
}
