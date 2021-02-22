using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.doc.sheet_
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

		static public doc.Sheet Mesh(WorkbookPart workbookPart)
		{
			
			return new Sheet(workbookPart, GetFirstWorksheet(workbookPart) ) ;
		}


		static public Sheet GetFirst(string file)
		{
			var doc = nilnul.blob_.excel.DocX.GetWorkbookPart(file);
			return new nilnul.blob_.excel.doc.Sheet(doc.Workbook.WorkbookPart, doc.WorksheetParts.First().Worksheet

				);
		}


	}
}
