using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet.dwelled.cel.merged
{
	class GetMerged
	{
		static void Main(WorksheetPart worksheetPart, Worksheet worksheet, Cell cell)
		{
			
			List<MergeCells> mergedCellsList;

			if (worksheet.Elements<MergeCells>().Count() > 0)
			{


				if (worksheet.Elements<MergeCells>().Count() > 0)
				{
					MergeCells
						mergedCells = worksheet.Elements<MergeCells>().First();

					foreach (MergeCell mc in mergedCells)
					{
						if (mc.Reference.InnerText.StartsWith("A3"))
						{
							var theCell = worksheetPart.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().
							   Where(c => c.CellReference == "C3").FirstOrDefault();
							//userDetails.FirstName = GetCellValue(spreadSheetDocument, theCell);
						}

					}
				}
			}
		}
	}
}
