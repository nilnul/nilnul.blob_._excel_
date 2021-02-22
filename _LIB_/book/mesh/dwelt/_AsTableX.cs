using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using coordRow = nilnul.obj._matrix._coord_._row;

namespace nilnul.fs.excel.doc.sheet.dwelt
{
	static public class _AsTableX
	{

		/// <summary>
		/// column caption added.
		/// </summary>
		/// <param name="doc"></param>
		/// <param name="mapping">
		///	map some (not all) cols
		/// </param>
		/// <param name="threshold">whether the belt is header</param>
		/// <returns> table is empty if header not found </returns>
		static public DataTable GetTable(
			SpreadsheetDocument doc,
			nilnul.txt.rel_.bijection.Partial mapping
			,
			nilnul.num.RealI threshold
		)
		{

			var r = new System.Data.DataTable();


			var head = belt._recs._FindHeaderX.GetCols(doc, mapping, threshold);
			if (head.Item2 ==null)
			{
				return r;
			}

			r.Columns.AddRange(
				head.Item1.OrderBy(
					c => c.Key,
					nilnul.obj._matrix._coord_._col.val.Comp.Singleton
				).Select(
					t => t.Value
				).ToArray()

			);

			var sheetFirst = nilnul.fs.excel.doc.sheets.choose_._FirstX.GetWorksheet(doc);

			var rowBegin = head.Item2; // the previous of data.


			var rowEnd = nilnul.fs.excel.doc.sheet.dwelt._AsDiagX.Get(
				nilnul.fs.excel.doc.sheets.choose_._FirstX.GetWorksheet(doc)
			).end.row;


			for (nilnul.obj._matrix._coord_._row.ValI row = coordRow.val.convert_.Inc.Singleton.convert(rowBegin); nilnul.obj._matrix._coord_._row.val.comp.Re.Singleton.le(row, rowEnd); row = coordRow.val.convert_.Inc.Singleton.convert(row))
			{
				var cels = nilnul.fs.excel.doc.sheet.dwelt.row._CelsX.Enumerate(
					sheetFirst,
					row
				).Where(
						c => head.Item1.Keys.Contains(c.col, nilnul.obj._matrix._coord_._col.val.Eq.Singleton)
					).OrderBy(
						 c1 => c1.col, nilnul.obj._matrix._coord_._col.val.Comp.Singleton
					)
					.Select(
						ce => nilnul.fs.excel.doc.sheet.dwelt.row.cel.closure.val._GetX.GetVal(doc.WorkbookPart, sheetFirst, ce)
					).ToArray();
				if (
					cels.All(
						c=> c is null
						||
						nilnul.txt.be_.White.Singleton.be(c.ToString())
					)
				)
				{
					continue;
				}

				r.Rows.Add(
					cels
				);
			}
			return r;
		}

		/// <summary>
		/// column caption added
		/// </summary>
		/// <param name="file"></param>
		/// <param name="headers"></param>
		/// <param name="threshold"></param>
		/// <returns></returns>
		static public DataTable GetTable(
			string file,
			nilnul.txt.rel_.bijection.Partial headers,
			nilnul.num.RealI threshold

		)
		{

			using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
			{
				return GetTable(doc,headers,threshold);
			}


		}
	}
}