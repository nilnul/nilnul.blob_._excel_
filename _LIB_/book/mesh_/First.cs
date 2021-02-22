using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.book.mesh_
{
	/// <summary>
	/// 
	/// </summary>
	/// <remarks>
	/// <see cref="nameof(nilnul.fs.excel.doc.sheets.choose_._FirstX)"/>
	/// </remarks>
	static public class _FirstX
	{
		static public Worksheet GetFirstWorksheet(string file)
		{
			var doc = nilnul.blob_.excel.DocX.GetWorkbookPart(file);
			return doc.WorksheetParts.First().Worksheet;
		}

		static public Worksheet GetFirstWorksheet(WorkbookPart doc)
		{
			
			return doc.WorksheetParts.First().Worksheet;
		}

		static public book.Mesh Mesh(SpreadsheetDocument workbookPart)
		{
			
			return new book.Mesh(workbookPart, GetFirstWorksheetPart(workbookPart) ) ;
		}

		private static WorksheetPart GetFirstWorksheetPart(SpreadsheetDocument doc)
		{
			return doc.WorkbookPart.WorksheetParts.First();
		}

		static public book.Mesh GetFirst(string file)
		{
			var doc =excel._BookX.Open(file);
			return new nilnul.blob_.excel.book.Mesh(
				doc, doc.WorkbookPart.WorksheetParts.First()

				);
		}


	}
}
