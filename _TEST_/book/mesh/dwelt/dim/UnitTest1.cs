using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Diagnostics;
using DocumentFormat.OpenXml.Packaging;
using System.Linq;
using System.Linq.Expressions;
using System.Collections.Generic;
using nilnul.obj.str.be_;

namespace nilnul.fs.excel.TEST.doc.sheet.dwelt.dim
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

				new nilnul.fs._address.doc_.Dotted("170626XiangFields.mapping", "xlsx")

			);

			var cols = new nilnul.txt.Set();

			using (FileStream fs = new FileStream(xlsx.ToString(), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))

			using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
			{

				var sheetFirst = nilnul.fs.excel.doc.sheets.choose_._FirstX.GetWorksheet(doc);

				var dim = nilnul.fs.excel.doc.sheet.dwelt._AsDiagX.Get( sheetFirst);
				Debug.WriteLine(dim);
			}


		}
	}
}
