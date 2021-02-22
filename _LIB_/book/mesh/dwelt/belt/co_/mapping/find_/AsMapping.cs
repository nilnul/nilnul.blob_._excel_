using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using nilnul.fs.excel.doc._sheet._coord_._row;
using DocumentFormat.OpenXml.Spreadsheet;

namespace nilnul.fs.excel.doc.sheet.dwelt.closures.belts.find
{
	/// <summary>
	/// given a template, find the headers
	/// </summary>
	static public class _AsMappingX
	{
		/// <summary>
		/// get the relation between src and target by enumerating cols.
		/// This is a total surjective relation. may be polygomy; it also can be polyandry as there may be white keys.
		/// </summary>
		/// <param name="doc"></param>
		/// <param name="worksheet"></param>
		/// <param name="row"></param>
		/// <returns></returns>
		public static IEnumerable<nilnul.txt.Duo> GetTxtDuoS(
			SpreadsheetDocument doc
			,
			Worksheet worksheet
			,
			nilnul.obj._matrix._coord_._row.ValI row
		)
		{

			var dim = nilnul.fs.excel.doc.sheet.dwelt._AsDiagX.Get(worksheet);

			var belts = nilnul.fs.excel.doc.sheet.dwelt.closures.Belts.Enumerate(
				worksheet
			);


			// get the belt from row

			var belt = belts.FirstOrDefault(b => nilnul.obj._matrix._coord_._row.val.Eq.Singleton.Equals(
				b.bounding.rowRange.lower
				,
				row
			));

			//get the next belt, which is the target
			var targets = belts.FirstOrDefault(
				b => b.bounding.rowRange.contain(
					nilnul.obj._matrix._coord_._row.val.convert_.Inc.Singleton.convert(
						belt.bounding.rowRange.upper
					)
				)
			);

			for (
				nilnul.obj._matrix._coord_._col.ValI i = dim.colRange.lower;
				nilnul.obj._matrix._coord_._col.val.comp.Re.Singleton.le(i, dim.colRange.upper);
				i = nilnul.obj._matrix._coord_._col.val.convert_.Inc.Singleton.convert(i)
			)
			{
				yield return new txt.Duo(

					nilnul.txt.op_.PurgeWhite.Singleton.op(
						nilnul.fs.excel.doc.sheet.dwelt.closures.belt.col._TxtX.GetTxt(
							doc.WorkbookPart
							,
							worksheet,
							belt,
							i
						)
					)
					,
					nilnul.fs.excel.doc.sheet.dwelt.closures.belt.col._TxtX.GetTxt(
						doc.WorkbookPart
						,
						worksheet,
						targets,
						i
					)
				);
			}
		}


		/// <summary>
		/// white keys are ignored, together with their values (if there is no key but value, it's ignored); white values are not mapped, leave some widowers. 
		/// </summary>
		/// <param name="doc"></param>
		/// <param name="worksheet"></param>
		/// <param name="row"></param>
		/// <returns></returns>
		public static nilnul.txt.rel_.bijection.Partial Get(
			SpreadsheetDocument doc
			,
			Worksheet worksheet
			,
			nilnul.obj._matrix._coord_._row.ValI row
		)
		{
			return nilnul.txt.rel_.bijection.Partial.Create_valWhiteNotMapped(
				new nilnul.txt.Dict(    //vow keys are distinct
					GetTxtDuoS(doc, worksheet, row).Where(
						x => nilnul.txt.be_.NonWhite.Singleton.be(x.Item1)  //white keys are ignored
					)
				)
			);
		}

		/// <summary>
		/// white keys are ignored, together with their values (if there is no key but value, it's ignored); white values are not mapped, leave some widowers. 
		/// </summary>
		public static nilnul.txt.rel_.bijection.Partial Get(nilnul.fs.address_.SpearI xlsx, nilnul.obj._matrix._coord_._row.ValI row)
		{
			using (FileStream fs = new FileStream(xlsx.ToString(), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
			{
				return Get(doc, nilnul.fs.excel.doc.sheets.choose_._FirstX.GetWorksheet(doc), row);
			}
		}
		/// <summary>
		/// white keys are ignored, together with their values (if there is no key but value, it's ignored); white values are not mapped, leave some widowers. 
		/// </summary>
		static public nilnul.txt.rel_.bijection.Partial Get(

			string xlsxPath
			,
			nilnul.obj._matrix._coord_._row.ValI row

			)
		{
			return Get(nilnul.fs.address_.Spear.Parse(xlsxPath), row);
		}

		/// <summary>
		/// white keys are ignored, together with their values (if there is no key but value, it's ignored); white values are not mapped, leave some widowers. 
		/// </summary>
		static public nilnul.txt.rel_.bijection.Partial Get(

			string xlsxPath
			,
			int row

			)
		{
			return Get(xlsxPath, new nilnul.fs.excel.doc._sheet._coord_._row.Val(row));
		}
		/// <summary>
		/// white keys are ignored, together with their values (if there is no key but value, it's ignored); white values are not mapped, leave some widowers. 
		/// </summary>
		public static nilnul.txt.rel_.bijection.Partial Get(byte[] xlsx, nilnul.obj._matrix._coord_._row.ValI row)
		{
			using (Stream fs = new MemoryStream(xlsx, false))
			using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
			{
				return Get(doc, nilnul.fs.excel.doc.sheets.choose_._FirstX.GetWorksheet(doc), row);
			}
		}
		/// <summary>
		/// white keys are ignored, together with their values (if there is no key but value, it's ignored); white values are not mapped, leave some widowers. 
		/// </summary>
		public static nilnul.txt.rel_.bijection.Partial Get(byte[] xlsx, int row)
		{
			return Get(xlsx, new nilnul.fs.excel.doc._sheet._coord_._row.Val(row));
		}
		/// <summary>
		/// white keys are ignored, together with their values (if there is no key but value, it's ignored); white values are not mapped, leave some widowers. 
		/// </summary>
		/// <param name="xlsx"></param>
		/// <param name="row"></param>
		/// <returns></returns>
		public static nilnul.txt.rel_.bijection.Partial Get(Stream xlsx, nilnul.obj._matrix._coord_._row.ValI row)
		{
			using (SpreadsheetDocument doc = SpreadsheetDocument.Open(xlsx, false))
			{
				return Get(doc, nilnul.fs.excel.doc.sheets.choose_._FirstX.GetWorksheet(doc), row);
			}
		}
		/// <summary>
		/// white keys are ignored, together with their values (if there is no key but value, it's ignored); white values are not mapped, leave some widowers. 
		/// </summary>
		/// <param name="xlsx"></param>
		/// <param name="row"></param>
		/// <returns></returns>
		public static nilnul.txt.rel_.bijection.Partial Get(Stream xlsx, int row)
		{
			return Get(xlsx, new nilnul.fs.excel.doc._sheet._coord_._row.Val(row));
		}
	}
}