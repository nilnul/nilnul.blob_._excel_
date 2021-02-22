using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet.dwelt.row.cel.block
{
	/// <summary>
	/// 
	/// </summary>
	/// 
	[Obsolete(nameof(cel._ValContainerX))]
	static public class _ValContainerX
	{
		static public nilnul.obj.matrix.Block ByMergeCellReference(WorksheetPart wsPart, DocumentFormat.OpenXml.Spreadsheet.Cell cell_ofSameSheet)
		{
			return ByMergeCellReference(wsPart.Worksheet,cell_ofSameSheet);

			//return MergeCel.Cast(
			//	sheet.mergeCellS.Get.Enumerate(wsPart)
			//	.FirstOrDefault(
			//		x => MergeCel.Cast(x).conain(
			//			row.cel._key.Ret.FroCell(cell_ofSameSheet)
			//		)
			//	)
			//);
		}

		/// <summary>
		/// supposing at most one MergeCell containing the cell
		/// </summary>
		/// <param name="wsPart"></param>
		/// <param name="cell_ofSameSheet"></param>
		/// <returns></returns>
		static public nilnul.obj.matrix.Block ByMergeCellReference(Worksheet wsPart, DocumentFormat.OpenXml.Spreadsheet.Cell cell_ofSameSheet)
		{
			var mergeCells = sheet.mergeCellS.Get.Enumerate(wsPart).Where(
				x => MergeCel.Cast(x).conain(
					row.cel._key.Ret.FroCell(cell_ofSameSheet)
				)
			
			);
			if (mergeCells.Any())
			{
				return MergeCel.Cast(
								mergeCells
								.First(
								)
							);
			}
			else
			{
				return  cel._AsBlockX.AsBlock(cell_ofSameSheet);
			}


			
		}


	}
}
