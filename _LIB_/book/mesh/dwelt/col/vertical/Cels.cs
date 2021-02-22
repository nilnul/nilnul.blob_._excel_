using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;

namespace nilnul.blob_.excel.doc.sheet.dwelt.col.vertical
{
	static public class _HeadingX
	{

		


		public static string Nulable(SpreadsheetDocument document,  Worksheet worksheet, string columnName)
		{

			// Get the cells in the specified column and order them by row.
			IEnumerable<Cell> cells =cels_._ReifiedX.Cels(worksheet,columnName);

			if (cells.Count() == 0)
			{
				// The specified column does not exist.
				return null;
			}

			// Get the first cell in the column.
			Cell headCell = cells.First();

			return cel._ValX.GetTxt(document, headCell);

		}
	}
}