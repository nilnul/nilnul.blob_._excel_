using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using nilnul.blob_.excel.doc;
using nilnul.fs.excel.doc.sheet.dwelt._cel;
using nilnul.obj.mesh._cel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.book.mesh.cel.val_.typed
{
	static public class _GetX
	{
		static public object GetObj(WorkbookPart workbookPart, SheetData sheetData, blob_.excel.doc.sheet._cel.Coord coord)
		{
			//undone
			var c = doc.sheet._CelX.ByCoord(
				sheetData, coord
			);

			return GetObj(workbookPart, c);

		}

		static public object GetObj(WorkbookPart workbookPart, Worksheet sheetData, blob_.excel.doc.sheet._cel.Coord coord)
		{
			return GetObj(
				workbookPart
				,
				excel.doc.sheet.DweltX.GetSheetData(sheetData)
				, coord
			);

		}
		static public object GetObj(WorkbookPart workbookPart, WorksheetPart sheetData, blob_.excel.doc.sheet._cel.Coord coord)
		{
			

			return GetObj(workbookPart, sheetData.Worksheet,coord);

		}
		public static object GetObj(blob_.excel.doc.Sheet dwelt, blob_.excel.doc.sheet._cel.Coord x)
		{
			return GetObj(dwelt.workbookPart, dwelt.worksheet, x);
		}

		public static object GetObj(blob_.excel.book.Mesh dwelt, blob_.excel.doc.sheet._cel.Coord x)
		{
			return GetObj(dwelt.workbookPart, dwelt.worksheetPart, x);
		}


		public static object GetObj(blob_.excel.doc.Sheet dwelt, Coord x)
		{
			return GetObj(dwelt, new blob_.excel.doc.sheet._cel.Coord(x));
		}

		public static object GetObj(blob_.excel.doc.Sheet dwelt, int x, int y)
		{
			return GetObj(dwelt,  blob_.excel.doc.sheet._cel.Coord.OvColRowNilBased(x,y));
		}

		public static object GetObj_rowOneBased(blob_.excel.doc.Sheet dwelt, int col, int row)
		{
			return GetObj(dwelt,  blob_.excel.doc.sheet._cel.Coord.OvColRowNilBased(col,row-1));
		}

		public static object GetObj_rowOneBased(blob_.excel.book.Mesh dwelt, int col, int row)
		{
			return GetObj(dwelt,  blob_.excel.doc.sheet._cel.Coord.OvColRowNilBased(col,row-1));
		}






		/*
		 
			 OpenXML documentation on msdn


The value of the DataType property is null for numeric and date types. It contains the value CellValues.SharedString for strings, and CellValues.Boolean for Boolean values.
*/
		static public object GetObj(WorkbookPart workbookPart, Cell cell)
		{
			if (cell == null)
			{
				return null;
			}
			if (cell.CellValue == null)
			{
				return null;
			}

			string txt = cell.CellValue.InnerText;

			object r = txt;

			//if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
			//{
			//	return workbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(int.Parse(value.ToString())).InnerText;
			//}

			if (cell.DataType != null)
			{
				switch (cell.DataType.Value)
				{
					case CellValues.SharedString:
						int ssid = int.Parse(cell.CellValue.Text);
						r = workbookPart.SharedStringTablePart.SharedStringTable.ElementAt(ssid).InnerText;
						break;

					case CellValues.Boolean:
						switch (txt)
						{
							case "0":
								r = false;
								break;
							default:
								r = true;
								break;
						}
						break;

					case CellValues.Date:

						r = DateTime.Parse(txt);

						break;
					case CellValues.Error
					:
						break;
					case CellValues.InlineString:
						break;
					case CellValues.Number:
						r = double.Parse(txt);
						break;

					case CellValues.String:
						break;
				}
			}
			else
			{/*
				OpenXML documentation on msdn

The value of the DataType property is null for numeric and date types.It contains the value CellValues.SharedString for strings, and CellValues.Boolean for Boolean values.
*/
				var styleIndex1 = cell.StyleIndex;
				if ((styleIndex1 is null))
				{
					double numVal;
					if (double.TryParse(cell.InnerText, out numVal))
					{
						r = numVal;

					}
				}
				else
				{
					int styleIndex = (int)cell.StyleIndex.Value;

					CellFormat cellFormat = (CellFormat)workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ElementAt(styleIndex);

					uint formatId = cellFormat.NumberFormatId.Value;

					if (formatId == (uint)NumDateFormatId.DateShort || formatId == (uint)NumDateFormatId.DateLong)
					{
						double oaDate;
						if (double.TryParse(txt, out oaDate))
						{
							r = DateTime.FromOADate(oaDate).ToShortDateString();
						}
					}
					else
					{
						double numVal;
						if (double.TryParse(cell.InnerText, out numVal))
						{
							r = numVal;

						}

					}

				}

			}

			return r;
		}

		static public object Obj(book.Mesh workbookPart, Cell cell)
		{
			return GetObj(workbookPart.workbookPart,cell);
		}
		public static object Obj(Cel cel)
		{
			return Obj(cel.mesh, cel.cell);
		}
	}
}

