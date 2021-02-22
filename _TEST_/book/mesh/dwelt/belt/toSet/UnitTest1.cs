using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Diagnostics;
using DocumentFormat.OpenXml.Packaging;
using System.Linq;
using System.Linq.Expressions;
using System.Collections.Generic;
using nilnul.obj.str.be_;

namespace nilnul.fs.excel.TEST.doc.sheet.dwelt.belt.toSet
{
	[TestClass]
	public class UnitTest1
	{
		[TestMethod]
		public void TestMethod1()
		{
			var processBase = nilnul.win.process_.dotNet.Addresses.container_AppDomainBase;

			var prjBase = nilnul.fs.address_.volRoute_.container.convert_.UpN._Eval(processBase, 2u);

			var dataContainer = nilnul.fs.address_.volRoute_.container.convert_.join_._DirX.Eval_ofContainerDst(
				prjBase, "_data"
			);


			var xlsx = nilnul.fs.address_.volRoute_.container.to_.Element.JoinDoc(dataContainer,

				new nilnul.fs._address.doc_.Dotted("170626XiangFields", "xlsx")

			);

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
					new nilnul.fs.excel.doc._sheet._coord_._row.Val(3u)
				));






				//Debug.WriteLine(
				//	nilnul.txts.accumulate_.Commaed.Singleton.accumulate(
				//		nilnul.fs.excel.doc.sheet.dwelt.closures.belt.cols._TxtX.GetTxts(
				//			doc.WorkbookPart,
				//			sheetFirst
				//			, belt

				//		//	new excel.doc.sheet.dwelt.closures.Belt(
				//		//		 excel.doc.sheet.Dwelt._Create(
				//		//			 doc,sheetFirst
				//		//		)
				//		//		,
				//		//		belt
				//		//	)

				//		).Where(x => !string.IsNullOrWhiteSpace(x))
				//	)
				//);


				var bag = new nilnul.txt.Bag(
					nilnul.fs.excel.doc.sheet.dwelt.closures.belt.cols._TxtX.GetTxts(
							doc.WorkbookPart,
							sheetFirst
							, belt

						//	new excel.doc.sheet.dwelt.closures.Belt(
						//		 excel.doc.sheet.Dwelt._Create(
						//			 doc,sheetFirst
						//		)
						//		,
						//		belt
						//	)

						).Select(x => nilnul.txt.op_.PurgeWhite.Singleton.op(x)).Where(x => !string.IsNullOrWhiteSpace(x))
				);//.toSet();

				var non = new nilnul.obj.bag.be_.nonPlural.vow.Ed<string, nilnul.txt.Eq>(bag);

				var setOfPresetColumns = non.toSet();

				Debug.WriteLine(
					non.toSet()
				);


				cols = setOfPresetColumns;
			}


			using (FileStream fs = new FileStream(xlsx.ToString(), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
			{
				var sheetFirst = nilnul.fs.excel.doc.sheets.choose_._FirstX.GetWorksheet(doc);

				var belts = nilnul.fs.excel.doc.sheet.dwelt.closures.Belts.Enumerate(
					sheetFirst
				);

				var threshold = new nilnul.num.real_.GoldenRatioConjugateComplement();


				var enumerator = belts.GetEnumerator();

				nilnul.obj._matrix._coord_._row.ValI rowOfDataBase_pre = null;

				Dictionary<nilnul.obj._matrix._coord_._col.ValI, string> colsFound = new Dictionary<obj._matrix._coord_._col.ValI, string>(
					nilnul.obj._matrix._coord_._col.val.Eq.Singleton
				);

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
						cols
					);

					var bagOfColHeader = nilnul.obj.bag.op_._PackX.Op<string, nilnul.txt.Eq>(bag);



					var colCount = bag.Keys.Count;

					var proportion = nilnul.num.Quotient.CreateByDivide(colCount, cols.Count);

					if (
						nilnul.num.real.Comparer.Decider.Singleton.ge(
							proportion,
							threshold
						)
					)
					{

						new nilnul.obj.bag.be_.nonPlural.Vow("some header appears plural times").vow(bag);



						rowOfDataBase_pre = enumerator.Current.bounding.rowRange.upper;

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

							;


						break;

					}


				}

				Debug.WriteLine(
					rowOfDataBase_pre
				);

				Debug.WriteLine(
					nilnul.obj.DictX.Phrase(
						colsFound
					)
				);



			}





		}
	}
}
