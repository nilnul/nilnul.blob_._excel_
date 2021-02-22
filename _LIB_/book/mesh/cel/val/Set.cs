using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.num.duo;

namespace nilnul.fs.excel.doc.sheet.dwelt.cel.val
{
	public class _SetX
	{
		public static void Set(string docName, string sheetName, string val,
			uint rowIndex_zeroBased, string columnName)
		{
			// Open the document for editing.
			using (SpreadsheetDocument spreadSheet =
					 SpreadsheetDocument.Open(docName, true))
			{
				WorksheetPart worksheetPart =
					sheets.choose_.ByName.GetWorksheetPart(spreadSheet, sheetName);

				if (worksheetPart != null)
				{
					Cell cell = _CelX.Get(
						worksheetPart.Worksheet
						 , 
						new nilnul.fs.excel.doc._sheet._coord_._row.Val((uint)rowIndex_zeroBased)
						,
						nilnul.fs.excel.doc._sheet._coord_._col.Val.Parse_bigEndian(  columnName)
					);

					cell.CellValue = new CellValue(val);
					cell.DataType =
						new EnumValue<CellValues>(CellValues.Number);

					// Save the worksheet.
					worksheetPart.Worksheet.Save();
				}
			}

		}





	}
}
