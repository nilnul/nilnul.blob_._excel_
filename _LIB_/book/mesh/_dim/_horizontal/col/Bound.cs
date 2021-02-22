using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.obj._matrix._coord_._col;

namespace nilnul.fs.excel.doc.sheet._dim._horizontal.col
{
	public class Bound:nilnul.obj._matrix._coord_._col.val.Range
	{
		private Bound(ValI row1, ValI row2) : base(row1, row2)
		{
		}

		public Bound( 
			Val a, Val b
		):base(a,b)
		{

		}

		static public Bound CreateSingleton(Val a) {
			return new Bound(a,a);
		}

		
	}
}
