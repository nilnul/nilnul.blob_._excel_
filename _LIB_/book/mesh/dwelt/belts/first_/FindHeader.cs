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

namespace nilnul.fs.excel.doc.sheet.dwelt.belt._recs
{
	public class _FindHeaderX
	{
		/// <summary>
		/// 
		/// </summary>
		/// <param name="docPath"></param>
		/// <param name="headers"></param>
		/// <param name="threshold"></param>
		/// <returns>
		///  row is null if no header row is found.
		///  row is the previous row to data if header is found
		///  
		/// </returns>
		static public Tuple<
			Dictionary<nilnul.obj._matrix._coord_._col.ValI, string>
			, nilnul.obj._matrix._coord_._row.ValI
		> FindHeaders(
			string docPath,
			nilnul.txt.Set headers,
			nilnul.num.RealI threshold
		)
		{
			var xlsx = nilnul.fs.address_.Spear.Parse(docPath);

			Dictionary<nilnul.obj._matrix._coord_._col.ValI, string> colsFound = new Dictionary<obj._matrix._coord_._col.ValI, string>(
				nilnul.obj._matrix._coord_._col.val.Eq.Singleton
			);


			using (FileStream fs = new FileStream(xlsx.ToString(), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
			{
				return FindHeaders(doc,headers,threshold);

			}
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
		/// <param name="headers"></param>
		/// <param name="threshold"></param>
		/// <returns>row null if not found; at this time the dict is empty</returns>
		static public Tuple<
			Dictionary<nilnul.obj._matrix._coord_._col.ValI, string>
			, nilnul.obj._matrix._coord_._row.ValI
		> FindHeaders(
			SpreadsheetDocument doc,
			nilnul.txt.Set headers,
			nilnul.num.RealI threshold
		)
		{

			Dictionary<nilnul.obj._matrix._coord_._col.ValI, string> colsFound = new Dictionary<obj._matrix._coord_._col.ValI, string>(
				nilnul.obj._matrix._coord_._col.val.Eq.Singleton
			);



			var sheetFirst = nilnul.fs.excel.doc.sheets.choose_._FirstX.GetWorksheet(doc);

			var belts = nilnul.fs.excel.doc.sheet.dwelt.closures.Belts.Enumerate(
				sheetFirst
			);

			//var threshold = new nilnul.num.real_.GoldenRatioLittle();

			var enumerator = belts.GetEnumerator();

			nilnul.obj._matrix._coord_._row.ValI rowOfDAtaBase_pre = null;

			while (enumerator.MoveNext())
			{


				//var b = enumerator.Current;

				var bag = new nilnul.txt.Bag(
					nilnul.fs.excel.doc.sheet.dwelt.closures.belt.cols._TxtX.GetTxts(
						doc.WorkbookPart,

							sheetFirst
							,

					enumerator.Current



						).Select(x => nilnul.txt.op_.PurgeWhite.Singleton.op(x)).Where(x => !string.IsNullOrWhiteSpace(x))
				);

				bag.keepOnly_ofFinite(
					headers
				);

				var bagOfColHeader = nilnul.obj.bag.op_._PackX.Op<string, nilnul.txt.Eq>(bag);

				var colCount = bag.Keys.Count;

				var proportion = nilnul.num.Quotient1.CreateByDivide(colCount, headers.Count);

				if (
					nilnul.num.real.comp.Re.Singleton.ge(
						new nilnul.num.real_.Quotient(
						proportion),
						threshold
					)
				)
				{

					new nilnul.obj.bag.be_.nonPlural.Vow("some header appears plural times").vow(bag);

					rowOfDAtaBase_pre = enumerator.Current.bounding.rowRange.upper;

					foreach (var item in nilnul.fs.excel.doc.sheet.dwelt.closures.belt.cols._TxtX.GetDict(

						doc.WorkbookPart,

							sheetFirst
							,

					enumerator.Current

					).Where(
						colVal => bagOfColHeader.Keys.Contains(colVal.Value)
					))
					{
						colsFound.Add(item.Key, item.Value);
					}

					break;
				}
			}

			return new Tuple<
				Dictionary<obj._matrix._coord_._col.ValI, string>
				,
				obj._matrix._coord_._row.ValI
			>(
				colsFound
				,
				rowOfDAtaBase_pre
			);


		}

		/// <summary>
		/// change caption to column name
		/// </summary>
		/// <param name="docPath"></param>
		/// <param name="headers"></param>
		/// <param name="threshold"></param>
		/// <returns></returns>
		public static Tuple<
			Dictionary<nilnul.obj._matrix._coord_._col.ValI, string>
			, nilnul.obj._matrix._coord_._row.ValI
		> FindHeaders(string docPath, Partial headers, RealI threshold)
		{

			var r = FindHeaders(
				docPath
				, headers.domain, threshold
			);

			//return r;

			var newDict = new Dictionary<nilnul.obj._matrix._coord_._col.ValI, string>();
			foreach (var item in r.Item1)
			{
				newDict.Add(item.Key, headers.bijection.ed[item.Value]);
			}

			return new Tuple<Dictionary<obj._matrix._coord_._col.ValI, string>, obj._matrix._coord_._row.ValI>(
				newDict
				,
				r.Item2
			);

		}


		public static Tuple<
			Dictionary<nilnul.obj._matrix._coord_._col.ValI, System.Data.DataColumn>
			, nilnul.obj._matrix._coord_._row.ValI
		> GetCols(string docPath, Partial headers, RealI threshold)
		{

			var r = FindHeaders(
				docPath
				, headers.domain, threshold
			);

			//return r;

			var newDict = new Dictionary<nilnul.obj._matrix._coord_._col.ValI, System.Data.DataColumn>();
			foreach (var item in r.Item1)
			{
				newDict.Add(item.Key, new System.Data.DataColumn(headers.bijection.ed[item.Value]) { Caption= item.Value});
			}

			return new Tuple<Dictionary<obj._matrix._coord_._col.ValI, System.Data.DataColumn>, obj._matrix._coord_._row.ValI>(
				newDict
				,
				r.Item2
			);

		}


		public static Tuple<
			Dictionary<nilnul.obj._matrix._coord_._col.ValI, string>
			, nilnul.obj._matrix._coord_._row.ValI
		> FindHeaders(
			SpreadsheetDocument docPath
			, Partial headers, RealI threshold)
		{
			var r = FindHeaders(
				docPath
				, headers.domain, threshold
			);

			var newDict = new Dictionary<nilnul.obj._matrix._coord_._col.ValI, string>();
			foreach (var item in r.Item1)
			{
				newDict.Add(item.Key, headers.bijection.ed[item.Value]);
			}

			return new Tuple<Dictionary<obj._matrix._coord_._col.ValI, string>, obj._matrix._coord_._row.ValI>(
				newDict
				,
				r.Item2
			);

		}
		public static Tuple<
			Dictionary<nilnul.obj._matrix._coord_._col.ValI, DataColumn>
			, nilnul.obj._matrix._coord_._row.ValI
		> GetCols(
			SpreadsheetDocument docPath
			, Partial headers, RealI threshold)
		{
			var r = FindHeaders(
				docPath
				, headers.domain, threshold
			);

			var newDict = new Dictionary<nilnul.obj._matrix._coord_._col.ValI, System.Data.DataColumn>();
			foreach (var item in r.Item1)
			{
				newDict.Add(
					item.Key,
					new System.Data.DataColumn(
						headers.bijection.ed[item.Value]
						, typeof(object)
					) {
						Caption = item.Value
					}
				);
			}

			return new Tuple<Dictionary<obj._matrix._coord_._col.ValI, DataColumn>, obj._matrix._coord_._row.ValI>(
				newDict
				,
				r.Item2
			);

		}


	}
}
