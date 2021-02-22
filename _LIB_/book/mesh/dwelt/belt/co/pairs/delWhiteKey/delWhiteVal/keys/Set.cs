using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using nilnul.fs.excel.doc._sheet._coord_._row;
using DocumentFormat.OpenXml.Spreadsheet;

namespace nilnul.blob_.excel.doc.sheet.dwelt.belt.co.pairs.delWhiteKey.delWhiteVal.keys
{
	/// <summary>
	/// </summary>
	static public class _SetX
	{
		

		public static nilnul.txt.Set TxtSet(
			SpreadsheetDocument doc
			,
			Worksheet worksheet
			,
			nilnul.obj._matrix._coord_._row.ValI row
			,
			nilnul.obj._matrix._coord_._row.ValI rowFilter
		)
		{
			return new nilnul.txt.Set(
				delWhiteKey.delWhiteVal._KeysX.Keys(doc, worksheet, row,rowFilter)
			); 
		}




		/// <summary>
		/// </summary>
		public static nilnul.txt.Set  TxtSet(
			nilnul.fs.address_.SpearI xlsx
			,
			nilnul.obj._matrix._coord_._row.ValI row
			,
			nilnul.obj._matrix._coord_._row.ValI rowFilter
		)
		{
			using (FileStream fs = new FileStream(xlsx.ToString(), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
			{
				return TxtSet(doc, nilnul.fs.excel.doc.sheets.choose_._FirstX.GetWorksheet(doc), row,rowFilter);
			}
		}
		/// <summary>
		/// </summary>
		static public nilnul.txt.Set TxtSet(

			string xlsxPath
			,
			nilnul.obj._matrix._coord_._row.ValI row
			,
			nilnul.obj._matrix._coord_._row.ValI rowFilter

			)
		{
			return TxtSet(nilnul.fs.address_.Spear.Parse(xlsxPath), row,rowFilter);
		}

		/// <summary>
		/// white keys are ignored, together with their values (if there is no key but value, it's ignored); white values are not mapped, leave some widowers. 
		/// </summary>
		static public nilnul.txt.Set TxtSet(

			string xlsxPath
			,
			int row
			,
			int rowFilter

			)
		{
			return TxtSet(
				xlsxPath, 
				new nilnul.fs.excel.doc._sheet._coord_._row.Val(row), 
				new nilnul.fs.excel.doc._sheet._coord_._row.Val(rowFilter));
		}
		/// <summary>
		/// </summary>
		public static nilnul.txt.Set TxtSet(byte[] xlsx
			, nilnul.obj._matrix._coord_._row.ValI row
			, nilnul.obj._matrix._coord_._row.ValI rowFilter
			)
		{
			using (Stream fs = new MemoryStream(xlsx, false))
			using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
			{
				return TxtSet(doc, nilnul.fs.excel.doc.sheets.choose_._FirstX.GetWorksheet(doc), row,rowFilter);
			}
		}
		/// <summary>
		/// </summary>
		public static nilnul.txt.Set TxtSet(byte[] xlsx
			, int row
			, int rowFilter
		)
		{
			return TxtSet(
				xlsx
				, new nilnul.fs.excel.doc._sheet._coord_._row.Val(row)
				, new nilnul.fs.excel.doc._sheet._coord_._row.Val(rowFilter)
				);
		}
		/// <summary>
		/// </summary>
		/// <param name="xlsx"></param>
		/// <param name="row"></param>
		/// <returns></returns>
		public static nilnul.txt.Set TxtSet(Stream xlsx
			, nilnul.obj._matrix._coord_._row.ValI row
			, nilnul.obj._matrix._coord_._row.ValI rowFilter
		)
		{
			using (SpreadsheetDocument doc = SpreadsheetDocument.Open(xlsx, false))
			{
				return TxtSet(doc, nilnul.fs.excel.doc.sheets.choose_._FirstX.GetWorksheet(doc), row,rowFilter);
			}
		}
		/// <summary>
		/// </summary>
		/// <param name="xlsx"></param>
		/// <param name="row"></param>
		/// <returns></returns>
		public static nilnul.txt.Set TxtSet(Stream xlsx
			, int row
			, int rowFilter
		)
		{
			return TxtSet(
				xlsx
				, new nilnul.fs.excel.doc._sheet._coord_._row.Val(row)
				, new nilnul.fs.excel.doc._sheet._coord_._row.Val(rowFilter)
			);
		}
	}
}