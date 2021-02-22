using nilnul.fs.excel.doc.sheet.dwelt.closures_;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet.dwelt.closures.belts
{
	static public class _EachTxtsX
	{
		static public IEnumerable<string[,]> GetEachTxts(
			IEnumerable<Belt> belts
		) {
			return belts.Select(a=> belt._GetX.GetTxtMatrix(a));
		}

		static public IEnumerable<IEnumerable< string>> GetTxtEnumerable(
			IEnumerable<Belt> belts
		) {
			return  belts.Select(
				belt=> closures.belt._GetX.GetTxts(

					
					belt
					

				)
			);
		}

	}
}
