using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc._sheet._coord_._row.val
{
	public class _GrowX
	{
		static public Val Grow(Val row, nilnul.NumI num) {
			return  Val.CreateFro0Based(
				nilnul.num.combine_.Add.Singleton.combine(row.toNum(), num)
			);
		}
	}
}
