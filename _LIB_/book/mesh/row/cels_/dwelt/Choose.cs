using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.book.mesh.row.cels_.dwelt
{
	static public class _ChooseX
	{
		static public Cell Cel0nul(Row  row, nilnul.num.OrdI2 col) {
			return  row?.Elements<Cell>().Where(


				c=>				nilnul.num.ord.Eq2.Singleton.Equals(

					book.mesh._cel.Coord.CreateFroCell(c).col
					,
					col
				)

			).FirstOrDefault();
		}

		static public Cell Cel0nul(Row  row, int col) {
			return Cel0nul(row, new nilnul.num.Ord2(col) );
		}

		static public Cell Cel0nul_ofColOneBased(Row  row, int col) {
			return Cel0nul(row,  nilnul.num.ord_.OneBased.FroOneBased(col) );
		}


	}
}
