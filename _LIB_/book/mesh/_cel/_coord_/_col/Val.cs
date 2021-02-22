using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Threading.Tasks;
using nilnul.obj._matrix._coord_._col;

namespace nilnul.fs.excel.doc._sheet._coord_._col
{
//	[Obsolete(nameof(nilnul.fs.excel.doc._sheet._coord.Col))]
	public class Val
		:
		//nilnul.num_.positive.Ord
		//,
		nilnul.obj._matrix._coord_._col.val_.Positive
	{

		public Val(int index):base(new num_.Positive( index))
		{
		}

		public Val(ValI col):this( (int) ( col.toNum().toBigint()+1 )
			
			)
		{
		}

		public Val(BigInteger x) : base(x)
		{
		}

		public override string ToString()
		{
			return nilnul.character_.cha.sortie_.sown.bijectiveNumBase_.CapitalLetter.Singleton.phrase(
				(int)this.boxed.val
			);

			;
		}
		static public Val Parse_bigEndian(string name) {

			return new Val(
				(int)(
				nilnul.character_.cha.sortie_.sown.bijectiveNumBase_.CapitalLetter.Singleton.parse_bigEndian(
					name
					)
				)
			);
		}

	}
}
