using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.obj._matrix;

namespace nilnul.fs.excel.doc.sheet.dwelt.cel.val
{

	/// <summary>
	/// get the value of the cell itself. not considering wether it's in a mergeCell
	/// </summary>
	/// 
	[Obsolete(nameof(nilnul.blob_.excel.doc.sheet.dwelt.cel._ValX))]
	static public class _GetX
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

		[Obsolete("need workbookPart to get shared string",true)]
		internal static string GetTxt(SheetData sheetData, CoordI begin)
		{
			return GetTxt(sheetData, new _sheet.Coord(begin));

			//throw new NotImplementedException();
		}

		/// <summary>
		/// if the cell is in mergeCell, this doesn't return the value of the mergeCell; rather, this return the value of the cell itself, which is null when it's not the upperLeft
		/// </summary>
		/// <param name="workbookPart"></param>
		/// <param name="worksheet"></param>
		/// <param name="coord"></param>
		/// <returns></returns>
		/// 
		[Obsolete("the cell is in mergeCell, this doesn't return the value of the mergeCell", false)]
		static public string GetTxt(WorkbookPart workbookPart, Worksheet worksheet, nilnul.fs.excel.doc._sheet.Coord coord)
		{
			//undone
			var c = dwelt._CelX.Get(
				worksheet, coord
			);

			return GetTxt(workbookPart, c);

		}
		static public string GetTxt(WorkbookPart workbookPart, SheetData sheetData, nilnul.obj._matrix.CoordI coord)
		{
			//undone
			var c = dwelt._CelX.Get(
				sheetData,  coord
			);

			return GetTxt(workbookPart, c);

		}

		[Obsolete(nameof(typed._GetX),true)]

		static public object GetObj(WorkbookPart workbookPart, SheetData sheetData, nilnul.obj._matrix.CoordI coord)
		{
			//undone
			var c = dwelt._CelX.Get(
				sheetData,  coord
			);

			return GetObj(workbookPart, c);

		}


		[Obsolete(nameof(typed._GetX),true)]
		static public object GetObj(WorkbookPart workbookPart, Cell cell)
		{
			return typed._GetX.GetObj(workbookPart,cell);
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
								value = "false";// "FALSE";
								break;
							default:
								value = "true";// "TRUE";
								break;
						}
						break;
				}

			}

			return value;


		}
	}
}
