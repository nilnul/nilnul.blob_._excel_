using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/// <summary>
/// in the backend storage, namely the XML, MergeCels is stored under Element:"worksheet"
/// </summary>
namespace nilnul.blob_.excel.book.mesh
{
	public class _CelMergedS
	{
		static public MergeCells Collection(WorksheetPart wsPart)
		{

			 
			if (wsPart.Worksheet.Elements<MergeCells>().Count() > 0)
			{
				MergeCells mergeCells = wsPart.Worksheet.Elements<MergeCells>().First();
				return mergeCells;
			}

			return null;
		}

		static public MergeCells Collection(Worksheet sheet)
		{

			 
			if (sheet.Elements<MergeCells>().Count() > 0)
			{
				MergeCells mergeCells = sheet.Elements<MergeCells>().First();
				return mergeCells;
			}

			return null;
		}

		public static MergeCells Collection(Mesh mesh)
		{
			return Collection(mesh.worksheetPart);
		}

		static public IEnumerable<MergeCell> Enumerate(MergeCells mergeCellS)
		{
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

		static public IEnumerable<MergeCell> Enumerate(WorksheetPart wsPart)
		{
			var mergeCellS = Collection(wsPart);
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
		public static IEnumerable<MergeCell> Enumerate(Mesh mesh)
		{
			return Enumerate(mesh.worksheetPart);
		}


		static public IEnumerable<MergeCell> Enumerate(Worksheet sheet)
		{
			var mergeCellS = Collection(sheet);
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
