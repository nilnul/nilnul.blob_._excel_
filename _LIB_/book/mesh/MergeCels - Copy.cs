using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet.dwelt.closures_
{
	[Obsolete()]
	public class _MergeCels
	{
		static public MergeCells Eval(WorksheetPart wsPart)
		{

			 
			if (wsPart.Worksheet.Elements<MergeCells>().Count() > 0)
			{
				MergeCells mergeCells = wsPart.Worksheet.Elements<MergeCells>().First();
				return mergeCells;
			}

			return null;
		}

		

		static public MergeCells Eval(Worksheet sheet)
		{

			 
			if (sheet.Elements<MergeCells>().Count() > 0)
			{
				MergeCells mergeCells = sheet.Elements<MergeCells>().First();
				return mergeCells;
			}

			return null;
		}



		static public IEnumerable<MergeCell> Enumerate(WorksheetPart wsPart)
		{
			var mergeCellS = Eval(wsPart);
			if (mergeCellS==null)
			{
				yield break;
			}
			foreach (
				MergeCell mergeCell in mergeCellS
			)
			{
				yield return mergeCell;
			}
		}

		static public IEnumerable<MergeCell> Enumerate(Worksheet sheet)
		{
			var mergeCellS = Eval(sheet);
			if (mergeCellS==null)
			{
				yield break;
			}
			foreach (
				MergeCell mergeCell in mergeCellS
			)
			{
				yield return mergeCell;
			}
		}





	}
}
