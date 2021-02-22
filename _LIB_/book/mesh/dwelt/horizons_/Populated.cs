using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.book.mesh.dwelt.horizons_
{
	public class _PopulatedX
	{
		static public IEnumerable<Row> Get(DocumentFormat.OpenXml.Spreadsheet.SheetData sheetData)
		{
			return sheetData.Elements<Row>(); // Get the row IEnumerator

		}
		static public IEnumerable<Row> Get(Worksheet worksheet) {

			return Get(DweltX.GetSheetData(worksheet));

		}

		public static IEnumerable<Row> Get(doc.Sheet mesh)
		{
			return Get(mesh.worksheet);
		}
		public static IEnumerable<Row> Get(book.Mesh mesh)
		{
			return Get(mesh.worksheetPart.Worksheet);
		}
	}
}
