using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.num.ord.co_.colRow_;
using nilnul.obj._matrix._coord_._col;
using nilnul.obj._matrix._coord_._col.val;
using nilnul.obj._mesh._coord.col_;

namespace nilnul.blob_.excel.doc._sheet._coord.col
{
	public class Bound:
		nilnul.num.ord.Bound< nilnul.blob_.excel.doc._sheet._coord.Col>
		,
		IEnumerable< nilnul.obj._mesh._coord.col_.CaptialLetter>
	{
		public Bound(Col row1, Col row2) : base(row1, row2)
		{
		}

		

		public Bound(SameType<Col> val) : base(val)
		{
		}

		public Bound(Range colRange):this(colRange.lower,colRange.upper)
		{

		}

		public Bound(ValI lower, ValI upper):this(
			new Col(lower)
			,
			new Col(upper)
		)
		{
		}

		static public Bound CreateSingleton(Col a) {
			return new Bound(a,a);
		}

		public IEnumerator<CaptialLetter> GetEnumerator()
		{
			for (var i = this.en.Item1; i <= this.en.Item2;  i++)
			{
				yield return i;
			}

		}

		IEnumerator IEnumerable.GetEnumerator()
		{
			return GetEnumerator();
		}

		public nilnul.Num count {
			get {
				return new Num( this.upper.toNum().toBigint() - this.lower.toNum().toBigint() + 1);
			}
		}
		
	}
}
