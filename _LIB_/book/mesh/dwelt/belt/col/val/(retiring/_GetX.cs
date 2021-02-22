using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using nilnul.fs.excel.doc._sheet._coord_._row;
using DocumentFormat.OpenXml.Spreadsheet;

namespace nilnul.fs.excel.doc.sheet.dwelt.belt.col.val
{
	/// <summary>
	/// given a template, find the headers
	/// </summary>
	static public class _GetX
	{

		static public string Get(

			string xlsxPath
			,

			nilnul.obj._matrix._coord_._row.ValI row

			, nilnul.obj._matrix._coord_._col.ValI col

			)
		{

			return Get(
				nilnul.fs.address_.Spear.Parse(xlsxPath)

				, row
				, col
			);





		}

		public static string Get(
			nilnul.fs.address_.SpearI xlsx
			,
			nilnul.obj._matrix._coord_._row.ValI row
			, nilnul.obj._matrix._coord_._col.ValI col
		)
		{

			var cols = new nilnul.txt.Set();

			using (FileStream fs = new FileStream(xlsx.ToString(), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))

			using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
			{

				return Get(doc,nilnul.fs.excel.doc.sheets.choose_._FirstX.GetWorksheet(doc), row, col);




			}


		}

		public static string Get(
SpreadsheetDocument doc
			,
Worksheet worksheet
			,
			nilnul.obj._matrix._coord_._row.ValI row
			, nilnul.obj._matrix._coord_._col.ValI col
		)
		{
			var belts = nilnul.fs.excel.doc.sheet.dwelt.closures.Belts.Enumerate(
				worksheet
			);

			var belt = belts.FirstOrDefault(b => nilnul.obj._matrix._coord_._row.val.Eq.Singleton.Equals(
				b.bounding.rowRange.lower
				,
				row
			));

			var colVal = nilnul.fs.excel.doc.sheet.dwelt.closures.belt.col._TxtX.GetTxt(
				doc.WorkbookPart
				, worksheet, belt, col
			);
			return colVal;
		}
		public static string Get(
WorkbookPart doc
			,
Worksheet worksheet
			,
			nilnul.obj._matrix._coord_._row.ValI row
			, nilnul.obj._matrix._coord_._col.ValI col
		)
		{
			var belts = nilnul.fs.excel.doc.sheet.dwelt.closures.Belts.Enumerate(
				worksheet
			);

			var belt = belts.FirstOrDefault(b => nilnul.obj._matrix._coord_._row.val.Eq.Singleton.Equals(
				b.bounding.rowRange.lower
				,
				row
			));

			var colVal = nilnul.fs.excel.doc.sheet.dwelt.closures.belt.col._TxtX.GetTxt(
				doc
				, worksheet, belt, col
			);
			return colVal;
		}

	}

}
