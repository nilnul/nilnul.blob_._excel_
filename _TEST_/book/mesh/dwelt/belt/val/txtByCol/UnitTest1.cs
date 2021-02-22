using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Collections;
using System.Collections.Generic;

namespace nilnul.fs.excel.TEST.doc.sheet.dwelt.belt.val.txtByCol
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


			using (FileStream fs = new FileStream(xlsx.ToString(), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			
			using (SpreadsheetDocument doc= SpreadsheetDocument.Open(fs, false))
			{

				var sheetFirst = nilnul.fs.excel.doc.sheets.choose_._FirstX.GetWorksheet(doc);

				var belts = nilnul.fs.excel.doc.sheet.dwelt.closures.Belts.Enumerate(
					sheetFirst
				);

				foreach (var belt in belts)
				{

					Debug.WriteLine(
						nilnul.txts.accumulate_.Commaed.Singleton.accumulate(
							nilnul.fs.excel.doc.sheet.dwelt.closures.belt.cols._TxtX.GetTxts(
								doc.WorkbookPart,
								sheetFirst
								,belt

							//	new excel.doc.sheet.dwelt.closures.Belt(
							//		 excel.doc.sheet.Dwelt._Create(
							//			 doc,sheetFirst
							//		)
							//		,
							//		belt
							//	)
							
							).Where( x=>  !string.IsNullOrWhiteSpace(x))
						)
					);
				}
			}
		}
	}
}
