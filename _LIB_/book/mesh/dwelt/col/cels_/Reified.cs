using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using nilnul.fs.excel.doc._sheet._coord_._col;
using System;
using System.Collections.Generic;
using System.Linq;

namespace nilnul.blob_.excel.doc.sheet.dwelt.col.cels_
{
	static public class _ReifiedX
	{

		public static IEnumerable<Cell> Cels( Worksheet worksheet, string columnName)
		{


			// Get the cells in the specified column and order them by row.
			return  worksheet.Descendants<Cell>().Where(
				c => string.Compare(dwelt.cel._HorizontalIndexX.ColName(c.CellReference.Value), columnName, true) == 0
			);


		}

		public static IEnumerable<Cell> Ordered(  Worksheet worksheet, string columnName)
		{


			// Get the cells in the specified column and order them by row.
			return Cels(worksheet,columnName).OrderBy(
				r => cel._VertialIndexX.RowIndex(r.CellReference)
			);

			

		}

		public static IEnumerable<Cell> Ordered(Worksheet worksheet, Val col)
		{
			return Ordered(worksheet, col.ToString());
		}
	}
}