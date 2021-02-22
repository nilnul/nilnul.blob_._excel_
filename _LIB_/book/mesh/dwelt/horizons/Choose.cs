using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.obj._matrix._coord_._row;
using nilnul.num;

namespace nilnul.blob_.excel.book.mesh.dwelt.horizons
{

	/// <summary>
	/// </summary>
	static public class _ChooseX
	{
	

		public static Row Get(SheetData sheetData, nilnul.num.ord_.OneBased row)
		{
			return rows_._DweltX.Get(sheetData).Where(
				r =>  
					( r.RowIndex.Value) 
					==
					row.toOriginal().en
				
			
			).FirstOrDefault();
		}

		public static Row Get(Worksheet worksheet, nilnul.num.ord_.OneBased rowIndex)
		{
			return Get( DweltX.GetSheetData(worksheet),  rowIndex  );
		}

		public static Row Get(SheetData sheetData, OrdI2 row)
		{
			return Get((sheetData), new num.ord_.OneBased(row));

		}

		public static Row Get<TOrd>(SheetData sheetData, TOrd row)
			where TOrd:nilnul.num.OrdI2
		{
			return rows_._DweltX.Get(sheetData).Where(
				r => (
					 r.RowIndex.Value
					==
					row.toNum().toBigint().en+1
				)

			).FirstOrDefault();


			//throw new NotImplementedException();
		}

		public static Row Row0nul_ofRowOneBased(book.Mesh mesh, int row)
		{
			return rows_._DweltX.Get(mesh).Where(
			r =>
				(r.RowIndex.Value)
				==
				row


		).FirstOrDefault();

			///throw new NotImplementedException();
		}
	}


}
