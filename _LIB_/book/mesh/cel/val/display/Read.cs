using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet.dwelt.cel.val.display
{
	static public class _ReadX
	{
		static public string  GetDisplay(WorkbookPart workbookPart,Cell cell)
		{

			CellFormat cellFormat = (CellFormat)workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ElementAt( (int)(uint)cell.StyleIndex);


			string format = workbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats.Elements<DocumentFormat.OpenXml.Spreadsheet.NumberingFormat>()
			.Where(i => i.NumberFormatId.ToString() ==  cellFormat.NumberFormatId.ToString())
			.First().FormatCode;

			double d = double.Parse(cell.CellValue.InnerText);
			string val = d.ToString(format);



			return val;
		}

	}
}
