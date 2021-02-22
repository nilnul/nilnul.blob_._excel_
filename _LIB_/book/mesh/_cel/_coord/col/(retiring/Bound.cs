using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.blob_.excel.doc._sheet._coord;
using nilnul.obj._matrix._coord_._col;

namespace nilnul.fs.excel.doc._sheet._coord.col
{
	public class Bound:nilnul.obj._matrix._coord_._col.val.Range
	{
		public Bound(ValI row1, ValI row2) : base(row1, row2)
		{
		}

		public Bound( 
			Col a, Col b
		):base(a,b)
		{

		}

		static public Bound CreateSingleton(Col a) {
			return new Bound(a,a);
		}

		public nilnul.Num count {
			get {
				return new Num( this.upper.toNum().toBigint() - this.lower.toNum().toBigint() + 1);
			}
		}
		
	}
}
