using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.fs.excel.doc._sheet._coord_._row;

namespace nilnul.fs.excel.doc.sheet.dwelt.belt.cols
{
	/// <summary>
	/// given a template, find the headers
	/// </summary>
	static public class _AsHeadersX
	{

		static public nilnul.txt.Set GetPresetHeaders(
			
			string xlsxPath
			,
			nilnul.obj._matrix._coord_._row.ValI row
			

			)
		{

			var xlsx = nilnul.fs.address_.Spear.Parse(xlsxPath);

			var cols = new nilnul.txt.Set();

			using (FileStream fs = new FileStream(xlsx.ToString(), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))

			using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
			{

				var sheetFirst = nilnul.fs.excel.doc.sheets.choose_._FirstX.GetWorksheet(doc);

				var belts = nilnul.fs.excel.doc.sheet.dwelt.closures.Belts.Enumerate(
					sheetFirst
				);

				var belt = belts.FirstOrDefault(b => nilnul.obj._matrix._coord_._row.val.Eq.Singleton.Equals(
					b.bounding.rowRange.lower
					,
					row
				));

				var bag = new nilnul.txt.Bag(
					nilnul.fs.excel.doc.sheet.dwelt.closures.belt.cols._TxtX.GetTxts(
							doc.WorkbookPart,
							sheetFirst
							, belt

						).Select(x => nilnul.txt.op_.PurgeWhite.Singleton.op(x)).Where(x => !string.IsNullOrWhiteSpace(x))
				);//.toSet();

				var non = new nilnul.obj.bag.be_.nonPlural.vow.Ed<string, nilnul.txt.Eq>(bag);

				var setOfPresetColumns = non.toSet();

				


				cols = setOfPresetColumns;
			}
			return cols;
		}

		public static nilnul.txt.Set GetPresetHeaders(
			nilnul.fs.address_.SpearI xlsx
			,
			nilnul.obj._matrix._coord_._row.ValI  row
		)
		{

			var cols = new nilnul.txt.Set();

			using (FileStream fs = new FileStream(xlsx.ToString(), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))

			using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
			{

				var sheetFirst = nilnul.fs.excel.doc.sheets.choose_._FirstX.GetWorksheet(doc);

				var belts = nilnul.fs.excel.doc.sheet.dwelt.closures.Belts.Enumerate(
					sheetFirst
				);

				var belt = belts.FirstOrDefault(b => nilnul.obj._matrix._coord_._row.val.Eq.Singleton.Equals(
					b.bounding.rowRange.lower
					,
					row
				));




			

				var bag = new nilnul.txt.Bag(
					nilnul.fs.excel.doc.sheet.dwelt.closures.belt.cols._TxtX.GetTxts(
							doc.WorkbookPart,
							sheetFirst
							, belt

						).Select(x => nilnul.txt.op_.PurgeWhite.Singleton.op(x)).Where(x => !string.IsNullOrWhiteSpace(x))
				);//.toSet();

				var non = new nilnul.obj.bag.be_.nonPlural.vow.Ed<string, nilnul.txt.Eq>(bag);

				var setOfPresetColumns = non.toSet();

				


				cols = setOfPresetColumns;
			}


			return cols;
		}
	}
}
