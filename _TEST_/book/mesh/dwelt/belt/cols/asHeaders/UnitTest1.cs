using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Diagnostics;

namespace nilnul.fs.excel.TEST.doc.sheet.dwelt.belt.cols.asHeaders
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


			var cols = nilnul.fs.excel.doc.sheet.dwelt.belt.cols._AsHeadersX.GetPresetHeaders(
				xlsx
				,
				new nilnul.fs.excel.doc._sheet._coord_._row.Val(3u)
			);

			Debug.WriteLine(nilnul.txt.str.phrase_.Json.Singleton.phrase( cols));


		}
	}
}
