using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.part_.workbook
{
	static public class _Txts
	{
		static public SharedStringTablePart Get(WorkbookPart wbPart)
		{
		// For shared strings, look up the value in the
		// shared strings table.
		return wbPart.GetPartsOfType<SharedStringTablePart>()
			.FirstOrDefault();

		}


	}
}
