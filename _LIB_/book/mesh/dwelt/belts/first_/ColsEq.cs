using nilnul.fs.excel.doc.sheet.dwelt.closures;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Linq.Expressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace nilnul.blob_.excel.doc.sheet.dwelt.belts.first_
{

	public class ColsEqBunches
	{
		static public Belt Num(WorkbookPart workbookPart, Worksheet worksheet) {
			return Belts.EnumerateInBelt(workbookPart,worksheet).First(
				x=> nilnul.blob_.excel.doc.sheet.dwelt.belt._BunchesX.HorizontalBoundS(x).Count()
				==
				nilnul.blob_.excel.doc.sheet.dwelt._ColsX.Bound(worksheet).Count()
			);
		}
	}
}
