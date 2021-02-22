using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using nilnul.obj.matrix;
using nilnul.obj.matrix.block_;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.doc.sheet.dwelt.confine
{
	static public class _ValX
	{
		public static  string _Txt(
			WorkbookPart workbookPart
			,
			SheetData sheetData
			,
			nilnul.obj.matrix.BlockI _confine
		) {

			var diag = Diag.CreateFroBlock(_confine);

			return dwelt.cel._ValX.GetTxt(workbookPart, sheetData,   nilnul.obj.matrix.BlockIX.Begin( _confine));

		}

		public static string _Txt(
			WorkbookPart workbookPart
			,
			Worksheet worksheet
			,
			nilnul.obj.matrix.BlockI _confine
		)
		{
			return _Txt(workbookPart, nilnul.fs.excel.doc.sheet.DweltX.GetSheetData(worksheet), _confine);
		}

		internal static string _Txt(Belt belt, BlockI block)
		{
			return _Txt(
				belt.dwelt.sheet.workbookPart
				,
				belt.dwelt.sheet.worksheet
				,block
			);

		}
	}
}
