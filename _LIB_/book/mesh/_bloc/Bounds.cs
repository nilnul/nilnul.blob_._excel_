using nilnul.num.ord;
using nilnul.num.ord_;
using nilnul.num.ord_.oneBased_.bijective_;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.doc.sheet._bloc
{
	public class Bounds :
		nilnul.obj.mesh._bloc.Bounds1<nilnul.num.ord_.oneBased_.bijective_.UpperLetter, nilnul.num.ord_.OneBased>
	{


		public Bounds(Bound1<UpperLetter> colBound, Bound1<OneBased> rowBound) : base(colBound, rowBound)
		{
		}
		public Bounds(obj.mesh._bloc.Bounds bounds):this(
			new Bound1<UpperLetter>(
				new UpperLetter(
					bounds.colBound.lower
				)
				,
				new UpperLetter(
					bounds.colBound.upper
				)
			)
			,
			new Bound1<OneBased>(
				new OneBased(bounds.rowBound.lower)
				,
				new OneBased(bounds.rowBound.upper)
			)
		)
		{
		}
	}
}
