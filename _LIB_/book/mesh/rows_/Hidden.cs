using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;

namespace nilnul.blob_.excel.doc.sheet.dwelt.rows_
{
	public class Hidden
	{

		public static List<uint> Indexes(
			Worksheet workSheet

		)
		{
			return workSheet.Descendants<Row>()
				.Where((r) => r.Hidden != null && r.Hidden.Value)
				.Select(r => r.RowIndex.Value).ToList<uint>();
		}
	}
}
