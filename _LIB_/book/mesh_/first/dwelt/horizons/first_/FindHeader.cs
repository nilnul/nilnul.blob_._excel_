using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.num;
using nilnul.txt.rel_.bijection;
using System.Data;
using DocumentFormat.OpenXml.Spreadsheet;

namespace nilnul.blob_.excel.book.mesh_.first.dwelt.horizons.first_
{
	public class _ConfidentHeaderX
	{


		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
		/// <param name="headers"></param>
		/// <param name="threshold"></param>
		/// <returns>row null if not found; at this time the dict is empty</returns>
		static public Row FindHeaders(
			SpreadsheetDocument doc,
			nilnul.txt.Set headers,
			nilnul.num.RealI threshold
		)
		{




			var sheetFirst = nilnul.fs.excel.doc.sheets.choose_._FirstX.GetWorksheet(doc);

			var mesh = new blob_.excel.book.Mesh(doc, sheetFirst);
			var dwelt = new blob_.excel.book.mesh.Dwelt(mesh);



			var rows = blob_.excel.book.mesh.dwelt.Horizons.Row0nuls(
				sheetFirst
			);




			var enumerator1 = rows.GetEnumerator();





			while (enumerator1.MoveNext())
			{



				var set = excel.book.mesh.dwelt.horizon.cels_._NonblankX.TxtSet(

					dwelt,

					enumerator1.Current
				);

				var intersected = set.Intersect(headers, nilnul.txt.eq_.CaseInsensitive.Singleton);



				var colCount = intersected.Count();

				var proportion = nilnul.num.Quotient1.CreateByDivide(colCount, headers.Count);

				if (
					nilnul.num.real.comp.Re.Singleton.ge(
						new nilnul.num.real_.Quotient(
						proportion),
						threshold
					)
				)
				{


					return enumerator1.Current;



					break;
				}
			}
			return null;




		}







	}
}
