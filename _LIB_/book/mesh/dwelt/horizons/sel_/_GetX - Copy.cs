using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet.dwelt._rows
{


	public class _GetX
	{
		


		/// <summary>
		/// </summary>
		/// <param name="sheetData"></param>
		/// <returns></returns>
		static public IEnumerable<Row> Get_rowNullable( Worksheet worksheet)
		{
			var diag = dwelt._AsDiagX.Get(worksheet);

			var rows = dwelt.rows_.Populated.Get(worksheet).ToArray();

			foreach (var row in diag.rowRange)
			{
				yield return rows.Where(r => nilnul.obj._matrix._coord_._row.val.Eq.Singleton.Equals(row, new nilnul.fs.excel.doc._sheet._coord_._row.Val(r))).FirstOrDefault();
			}
		}

		static public nilnul.obj._matrix._coord_._row.val.Range RowIndexS(Worksheet worksheet) {
			return  dwelt._AsDiagX.Get(worksheet).rowRange;
		}

		public static IEnumerable<Row> Get_rowNullable(WorksheetPart worksheetPart)
		{
			return Get_rowNullable(worksheetPart.Worksheet);
			//throw new NotImplementedException();
		}


		

		public static IEnumerable<Row> Get_rowNullable(Workbook workbook, Sheet sheet)
		{
			return Get_rowNullable(

				doc.sheets.choose_.BySheet.GetWorksheet(workbook,sheet)
			);
			//throw new NotImplementedException();
		}


	}
}
