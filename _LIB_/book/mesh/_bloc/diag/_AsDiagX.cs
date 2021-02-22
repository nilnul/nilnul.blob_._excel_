using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.book.mesh._bloc.diag
{
	/// <summary>
	/// the end cell seems to have data.
	/// </summary>
	static public class _ParseX
	{

	
		static public nilnul.obj.mesh.Bloc Parse(string rangeReference)
		{
			var split = rangeReference.Split(':');
			if (split.Length == 1)
			{
				return obj.mesh.Bloc.CreateSingleton(
					nilnul.blob_.excel.book.mesh._cel.Coord.Parse(split[0])

					

				);
			}
			if (split.Length == 2)
			{
				return new obj.mesh.Bloc(
					nilnul.blob_.excel.book.mesh._cel.Coord.Parse(split[0])
					,
					nilnul.blob_.excel.book.mesh._cel.Coord.Parse(split[1])
				);
			}

			throw new UnexpectedReachException();

		}






		


	}
}
