using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.book.mesh.dwelt
{
	/// <summary>
	/// the end cell seems to have data.
	/// </summary>
	static public class _DiagX
	{

		static public nilnul.obj.matrix.block_.Diag Parse(string rangeReference)
		{
			var split = rangeReference.Split(':');
			if (split.Length == 1)
			{
				return obj.matrix.block_.Diag.CreateSingleton(
					nilnul.fs.excel.doc._sheet.Coord.Parse(split[0])

				);
			}
			if (split.Length == 2)
			{
				return new obj.matrix.block_.Diag(
					nilnul.fs.excel.doc._sheet.Coord.Parse(split[0])
					,
					nilnul.fs.excel.doc._sheet.Coord.Parse(split[1])
				);
			}

			throw new UnexpectedReachException();

		}

	



		static public nilnul.obj.matrix.block_.Diag Diag(SheetDimension dim)
		{

			return Parse(dim.Reference.Value);

		}
		static public nilnul.obj.mesh.Bloc Bloc(SheetDimension dim)
		{

			return mesh._bloc.diag._ParseX.Parse(dim.Reference.Value);

		}


		/// <summary>
		/// This element specifies the used range of the worksheet. It specifies the row and column bounds of used cells in the worksheet. This is optional and is not required. Used cells include cells with formulas, text content, and cell formatting. When an entire column is formatted, only the first cell in that column is considered used.
		/// </summary>
		/// <param name="worksheet"></param>
		/// <returns>
		/// the end cell is exclusive
		/// </returns>

		static public nilnul.obj.matrix.block_.Diag Get(Worksheet worksheet)
		{

			return Diag(worksheet.SheetDimension);

		}
		static public nilnul.obj.mesh.Bloc Bloc(Worksheet worksheet)
		{

			return Bloc(worksheet.SheetDimension);

		}


		public static nilnul.obj.matrix.block_.Diag Get(Mesh mesh)
		{
			return Get(mesh.worksheetPart.Worksheet);
		}



		


	}
}
