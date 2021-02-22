using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using System.Linq.Expressions;
using System.Data;

namespace nilnul.blob_._excel_._TEST_.book.mesh_.first.bloc.toTbl
{
	[TestClass]
	public class UnitTest1
	{
		[TestMethod]
		public void TestMethod1()
		{

			var path = @"C:\Users\me\Desktop\Book1.xlsx";

			var tbl = nilnul.blob_.excel.doc.sheet_.first.block.ToTblX.GetDataTbl(
				path,

				new obj.mesh._bloc.Bounds(
					 nilnul.num.ord_.oneBased_.bijective_.upper._BoundX.CreateBound(
						"I"
						,
						"J"
					),
					  nilnul.num.ord_.oneBased._BoundX.CreateGeneral(
						5
						,
						337
					)

				)
			);
			var rows = tbl.Rows.Cast<DataRow>().Select(x =>  (long.Parse( x.ItemArray[0].ToString()), x.ItemArray[1].ToString() )).ToArray();

			using (var db= new _l2s.DataClasses1DataContext(
				//@"data source=(localdb)\msSqlLocalDb;initial catalog=my;integrated security=True;MultipleActiveResultSets=True;App=EntityFramework"
))
			{
				foreach (var row in rows)
				{
					var currentR = db.Task1s.Where(t => t.id == row.Item1).Single();
					currentR.note = row.Item2;
				}

				db.SubmitChanges();
			}

		}
	}
}
