using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.doc.sheet.dwelt.cols_
{
	public class Hidden
	{
		/// <summary>
		/// start from 1?
		/// </summary>
		/// <param name="ws"></param>
		/// <returns></returns>
		public static List<uint> Indexes(Worksheet ws)
		{
			List<uint> itemList = new List<uint>();

			var cols = col._GroupsX.Groups(ws).
	  Where((c) => c.Hidden != null && c.Hidden.Value);

			foreach (Column item in cols)
			{
				for (uint i = item.Min.Value; i <= item.Max.Value; i++)
				{
					itemList.Add(i);
				}
			}

			return itemList;
		}
	}
}
