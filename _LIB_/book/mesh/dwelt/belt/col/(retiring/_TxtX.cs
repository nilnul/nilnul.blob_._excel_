using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.fs.excel.doc.sheet.dwelt.closures_;
using DocumentFormat.OpenXml.Packaging;
using nilnul.obj._matrix._coord_._col;
using nilnul.obj.matrix.block.set_.sed;

namespace nilnul.fs.excel.doc.sheet.dwelt.closures.belt.col
{
	//[Obsolete(nameof(blob_.excel.doc.sheet.dwelt.belt.col._ValX))]
	static public class _TxtX
	{
		static public string GetTxt(
			WorkbookPart workbookPart
			,

			SheetData sheetData
			,
			nilnul.obj.matrix.block.Set minBelt	//a collection of closures
			,
			nilnul.obj._matrix._coord_._col.ValI col
		) {
			var colClosures = minBelt.Where(
				x=> nilnul.obj.matrix.block.be_._IntersectColX.IntersectCol(x,col) 
			).OrderBy(
				x=>x
				,
				nilnul.obj.matrix.block.comp_.RowLower.Singleton
			);

			return string.Join("",
				colClosures.Select(
					
					block=> closure._TxtX.Get( workbookPart, sheetData,  block)
					
				)
			);

		}

		public static string GetTxt(
			WorkbookPart workbookPart,
			Worksheet worksheet, 
			nilnul.obj.matrix.block.Set minBelt
			,
			nilnul.obj._matrix._coord_._col.ValI col)
		{
			return GetTxt(  
				workbookPart,
				doc.sheet.DweltX.GetSheetData(worksheet) ,minBelt,col
				
			);

			//throw new NotImplementedException();
		}

		public static string GetTxt(WorkbookPart workbookPart, Worksheet sheetFirst, Bounding belt, ValI col)
		{

			return GetTxt(workbookPart, sheetFirst, belt.toOriginal(), col);
			//throw new NotImplementedException();
		}
	}
}
