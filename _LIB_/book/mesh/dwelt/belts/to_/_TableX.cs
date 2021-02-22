using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using coordRow = nilnul.obj._matrix._coord_._row;

namespace nilnul.blob_.excel.doc.sheet.dwelt.belts.to_
{
	static public class _TableX
	{
		
		/// <summary>
		/// column caption added.
		/// </summary>
		/// <param name="doc"></param>
		/// <param name="headers"></param>
		/// <param name="threshold"></param>
		/// <returns> table is empty if header not found </returns>
		static public DataTable

			GetTable1(
			SpreadsheetDocument doc,
			Worksheet sheet
		)
		{

			var r = new System.Data.DataTable();

			var headingBelt = nilnul.blob_.excel.doc.sheet.dwelt.belts.first_.ColsEqBunches.Num(doc.WorkbookPart, sheet);

			var txts = belt.cols.sel_._TxtX.GetTxts( headingBelt);

			r.Columns.AddRange(
				txts.Select(
					x=>new DataColumn( x)
				).ToArray()
			);

			var rowBegin = headingBelt.rows.en.Item2; // the previous of data.



			var rowEnd = nilnul.fs.excel.doc.sheet.dwelt._AsDiagX.Get(
				nilnul.fs.excel.doc.sheets.choose_._FirstX.GetWorksheet(doc)
			).end.row;

			for (
				nilnul.obj._matrix._coord_._row.ValI row = coordRow.val.convert_.Inc.Singleton.convert( new nilnul.obj._matrix._coord_._row.Val( rowBegin.toNum())); nilnul.obj._matrix._coord_._row.val.comp.Re.Singleton.le(row, rowEnd); row = coordRow.val.convert_.Inc.Singleton.convert(row))
			{

				r.Rows.Add(
					 nilnul.fs.excel.doc.sheet.dwelt.row._CelsX.Enumerate(
						sheet,
						row
					)
					.Select(
						ce => nilnul.fs.excel.doc.sheet.dwelt.row.cel.closure.val._GetX.GetVal(doc.WorkbookPart, sheet, ce)
					).ToArray()
				);
			}
			return r;
		}

		
	}
}