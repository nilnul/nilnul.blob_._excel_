using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.doc.sheet.dwelt.col
{
	static public class _GroupsX
	{
		

		/// <summary>
		/// this is infact columnGrp, though misnominated as Column
		/// </summary>
		/// <param name="ws"></param>
		/// <returns></returns>
		public static IEnumerable<Column> Groups(Worksheet ws)
		{

			return  ws.Descendants<Column>();
		}


	}
}
