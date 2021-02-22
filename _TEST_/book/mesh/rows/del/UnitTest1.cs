using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace nilnul.blob_._excel_._TEST_.book.mesh.rows.del
{
	[TestClass]
	public class UnitTest1
	{
		[TestMethod]
		public void TestMethod1()
		{

			var address = @"D:\170203\data\wyt.hbue.course_\webpage_\y20s(Git\exam\_score\regular.xlsx";
			var book = nilnul.blob_.excel._BookX.Open(address);

			//var workbookPart = nilnul.blob_.excel.DocX.GetWorkbookPart(address);
			//var firstSheet = blob_.excel.doc.sheet_._FirstX.GetFirstWorksheet(workbookPart);

			var mesh = blob_.excel.book.mesh_._FirstX.Mesh(book);

			var diag = blob_.excel.book.mesh.dwelt._DiagX.Get(mesh);

			var rows = excel.book.mesh.rows_._DweltX.Get(mesh.worksheetPart.Worksheet);
			foreach (var r in rows.ToArray())
			{

				var obj=blob_.excel.book.mesh.row.col.val_.typed._GetX.Obj_colNilBased(
					new excel.book.mesh.Major(mesh,r)
					
					,
					0
				);
				var txt = obj.ToString();
				if (txt.StartsWith("行政班"))
				{
					//remove the row

					blob_.excel.book.mesh.row._DropX.Drop(mesh, r);
				}
			}

			

			book.Save();
			book.Close();


		}
	}
}
