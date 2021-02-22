using nilnul.num.ord;
using nilnul.num.ord_;
using nilnul.num.ord_.oneBased_.bijective_;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.doc.sheet.bloc_._row
{
	public class Bounds : nilnul.obj.mesh._bloc.bounds_.Row1<nilnul.num.ord_.oneBased_.bijective_.UpperLetter, nilnul.num.ord_.OneBased>
	{

		public Bounds(Bound1<UpperLetter> colBound, OneBased rowBound) : base(colBound, rowBound)
		{
		}

		public Bounds(Bound1<UpperLetter> colBound, int v):this(colBound,new OneBased(v) )
		{
		}



		static public Bounds CreateHeader(Bound1<UpperLetter> colBound) {
			return new Bounds(colBound, 1);

		}
	}
}
