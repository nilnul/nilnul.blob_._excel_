using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using nilnul.fs.excel.doc.sheet.dwelt._cel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet.dwelt.cel.val.typed
{
	[Obsolete()]
	static public class _GetX
	{
		static public object GetObj(WorkbookPart workbookPart, SheetData sheetData, nilnul.obj._matrix.CoordI coord)
		{
			//undone
			var c = dwelt._CelX.Get(
				sheetData, coord
			);

			return GetObj(workbookPart, c);

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
					{	r = numVal;

					}

				}

			}

			return r;
		}
	}
}

