﻿using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using nilnul.obj.matrix.block_;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet.dwelt.closure.val
{
	[Obsolete()]
	static public class _GetX
	{
		public static  object Get(
			WorkbookPart workbookPart
			,
			SheetData sheetData
			,
			nilnul.obj.matrix.BlockI block
		) {

			var diag = Diag.CreateFroBlock(block);

			return dwelt.cel.val.typed._GetX.GetObj(workbookPart, sheetData,   nilnul.obj.matrix.BlockIX.Begin( block));

		}

		public static object Get(
			WorkbookPart workbookPart
			,
			Worksheet worksheet
			,
			nilnul.obj.matrix.BlockI block
		)
		{
			return Get(workbookPart, nilnul.fs.excel.doc.sheet.DweltX.GetSheetData(worksheet), block);
		}
	}
}
