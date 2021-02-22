using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet.dwelled.row.cel.val.get_
{
	public class Class1
	{
		public string GetCellValue(
			SpreadsheetDocument document
			,
			DocumentFormat.OpenXml.Spreadsheet.Cell cell
		)
		{
			SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
			string value = string.Empty;

			if (cell.CellValue != null)
			{
				value = cell.CellValue.InnerXml;
				if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
				{
					value = stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
				}
			}
			return value;
		}
	}
}
