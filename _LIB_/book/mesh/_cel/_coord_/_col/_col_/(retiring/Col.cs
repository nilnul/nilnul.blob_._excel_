using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet.dwelt.row.cel._key._ret
{
	[Obsolete()]
	public class Col
	{

		private int _index;

		/// <summary>
		/// zero based
		/// </summary>
		public int index
		{
			get { return _index; }
			set { _index = value; }
		}
		public Col(int index)
		{
			_index = index;
		}
		public override string ToString()
		{
			return nilnul.character.suc_.sown.bijectiveNumBase_.CapitalLetter.Singleton.phrase(_index+1);
				
				;
		}
		static public Col Parse_bigEndian(string name) {

			return new Col(
				(int)(
				nilnul.character.suc_.sown.bijectiveNumBase_.CapitalLetter.Singleton.parse_bigEndian(
					name
					)-1
				)
			);
		}

	}
}
