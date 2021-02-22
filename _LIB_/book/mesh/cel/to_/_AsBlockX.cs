using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet.dwelt.row.cel
{
	static public class _AsBlockX
	{
		static public nilnul.obj.matrix.Block AsBlock(this DocumentFormat.OpenXml.Spreadsheet.Cell cell)
		{

			return obj.matrix.Block.CreateSingleton(
				fs.excel.doc._sheet.Coord.CreateFroCell(cell)
			);
		}
	}
}
