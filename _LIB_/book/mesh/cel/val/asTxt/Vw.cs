using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.obj._matrix;
using nilnul.fs.excel.doc.sheet;

namespace nilnul.blob_.excel.doc.sheet.cel.val.asTxt
{

	/// <summary>
	/// get the value of the cell itself. not considering wether it's in a mergeCell
	/// </summary>
	static public class _VwX
	{
		static public string GetTxt(SpreadsheetDocument doc, Cell cell)
		{
			string value = cell.CellValue.InnerText;
			if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
			{
				return doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(int.Parse(value)).InnerText;
			}
			return value;
		}

		static public string GetTxt(WorkbookPart workbookPart, Cell cell)
		{
			if (cell == null)
			{
				return null;
			}
			if (cell.CellValue==null)
			{
				return null;
			}

			string value = cell.CellValue.InnerText;
			if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
			{
				return workbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(int.Parse(value)).InnerText;
			}
			return value;
		}


		static public string GetTxt(WorkbookPart workbookPart, SheetData sheetData, nilnul.obj._matrix.CoordI coord)
		{
			//undone
			var c = nilnul.fs.excel.doc.sheet. dwelt._CelX.Get(
				sheetData,  coord
			);

			return GetTxt(workbookPart, c);

		}

		static public string GetTxt<TCol,TRow>(
			WorkbookPart workbookPart,
			SheetData sheetData,
			nilnul.obj.mesh._cel.CoordI<TCol,TRow> coord
			)
			where TCol: nilnul.num.OrdI2
			where TRow: nilnul.num.OrdI2
		{
			//undone
			var c = nilnul.blob_.excel.doc.sheet. _CelX.ByCoord(
				sheetData,
				coord
			);

			return GetTxt(workbookPart, c);

		}

		public static string GetTxt(WorkbookPart dwelt, Worksheet sheet, obj.mesh._cel.Coord x)
		{
			return GetTxt(dwelt, sheet.GetSheetData(), x);
		}

		public static string GetTxt(Sheet dwelt, obj.mesh._cel.Coord x)
		{
			return GetTxt(dwelt.workbookPart,dwelt.worksheet, x);
		}


		static public string GetTxt(WorkbookPart workbookPart, SheetData sheetData, nilnul.blob_.excel.doc.sheet._cel.Coord coord)
		{
			//undone
			var c = nilnul.blob_.excel.doc.sheet. _CelX.ByCoord(
				sheetData,  coord
			);

			return GetTxt(workbookPart, c);

		}


	

		[Obsolete()]
		static public string GetVal(WorkbookPart workbookPart, Cell cell)
		{
			string value = cell.CellValue.InnerText;

			// If the cell represents an integer number, you are done. 
			// For dates, this code returns the serialized value that 
			// represents the date. The code handles strings and 
			// Booleans individually. For shared strings, the code 
			// looks up the corresponding value in the shared string 
			// table. For Booleans, the code converts the value into 
			// the words TRUE or FALSE.
			if (cell.DataType != null)
			{
				switch (cell.DataType.Value)
				{
					case CellValues.SharedString:

						// For shared strings, look up the value in the
						// shared strings table.
						value =
							workbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(int.Parse(value)).InnerText;

						// If the shared string table is missing, something 
						// is wrong. Return the index that is in
						// the cell. Otherwise, look up the correct text in 
						// the table.

						break;

					case CellValues.Boolean:
						switch (value)
						{
							case "0":
								value = "FALSE";
								break;
							default:
								value = "TRUE";
								break;
						}
						break;
				}

			}

			return value;


		}
	}
}
