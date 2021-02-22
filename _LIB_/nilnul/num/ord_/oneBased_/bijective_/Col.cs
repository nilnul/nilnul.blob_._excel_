using nilnul.num._ord_;
using nilnul.num_;
using nilnul.num_.bijective_;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.num.ord_.oneBased_.bijective_
//namespace nilnul.blob_.excel.doc.sheet._cel._coord

{
	[Obsolete(nameof(num.ord_.oneBased_.bijective_.UpperLetter))]


	public class Col : nilnul.num.ord_.OneBased
		,
		nilnul.num.OrdI1
		//,nilnul.num._ord_.ToBaseZeroI2
	{

		public Col(OrdI1 ord) : base(ord)
		{
		}

		protected Col(Positive1 val) : base(val)
		{
		}

		public Col(UpperLetters upperLetters):this( new Positive1(upperLetters))
		{
		}

		public override string ToString()
		{
			return new nilnul.num_.bijective_.UpperLetters(this.boxed).ToString();
		}

		static public Col Parse(string x) {
			return new Col(
nilnul.num_.bijective_.UpperLetters.Parse_bigEndian(x)
			);
		}
		
	}
}
