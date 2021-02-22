using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.num._ord_;
using nilnul.num_;
using nilnul.obj._matrix._coord_._col;

namespace nilnul.blob_.excel.doc.sheet._cel._coord
{
	public class Col
		:
		//nilnul.num_.positive.Ord
		//,
		nilnul.obj._mesh._coord.col_.CaptialLetter
		,
		nilnul.num.OrdI2
	{

		public Col(int index):base(new num_.Positive1( index))
		{
		}

		public Col(ValI col):this( (int) ( col.toNum().toBigint()+1 )
			
			)
		{
		}

		public Col(Positive1 val) : base(val)
		{
		}

		

		static public Col operator ++(Col col) {
			return new Col(
				new Positive1(
				col.val+1
				)
			);
		}
		static public Col Parse(string name) {

			return new Col(
				(int)(
				nilnul.character_.cha.sortie_.sown.bijectiveNumBase_.CapitalLetter.Singleton.parse_bigEndian(
					name
					)
				)
			);
		}

	

		NumI1 ToBaseZeroI3.toNum()
		{
			return new Num1(this.toNum());
		}
	}
}
