using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;

namespace nilnul.blob_.excel.book.mesh.dwelt
{
	/// <summary>
	/// divide the closures into belts
	/// </summary>
	public class Horizons
	{


		private Dwelt	 _dwelt;

		public Dwelt dwelt
		{
			get => _dwelt;
			set => _dwelt = value;
		}

		/// <summary>
		/// </summary>
		/// <param name="sheetData"></param>
		/// <returns></returns>
		static public IEnumerable<Row> Row0nuls(Worksheet worksheet)
		{
			var diag = book.mesh. dwelt._DiagX.Get(worksheet);

			var rows = book.mesh.dwelt.horizons_._PopulatedX.Get(worksheet).ToArray();

			foreach (var row in diag.rowRange)
			{
				yield return rows.Where(
					r => nilnul.obj._matrix._coord_._row.val.Eq.Singleton.Equals(row, new nilnul.fs.excel.doc._sheet._coord_._row.Val(r))).FirstOrDefault();
			}
		}

	

		public static IEnumerable<Row> Row0nuls(WorksheetPart worksheetPart)
		{
			return Row0nuls(worksheetPart.Worksheet);
			//throw new NotImplementedException();
		}




		public static IEnumerable<Row> Row0nuls(Workbook workbook, Sheet sheet)
		{
			return Row0nuls(

				book.meshs.choose_.BySheet.GetWorksheet(workbook, sheet)
			);
			//throw new NotImplementedException();
		}

	}
}
