using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.fs.excel.doc.sheet.dwelt.closures_;
using DocumentFormat.OpenXml.Packaging;

namespace nilnul.fs.excel.doc.sheet.dwelt.closures.belt.cols
{
	static public class _TxtX
	{

		static public IEnumerable< string > GetTxts(
			WorkbookPart workbookPart,
			SheetData sheetData
			,

			nilnul.obj.matrix.block.set_.sed.Bounding _belt
		)
		{

			return _belt.bounding.colRange.Select(
				c => GetTxt(
					workbookPart,
					sheetData
					,_belt.toOriginal()
					,c
					
				)
			);
		}
		static public Dictionary<nilnul.obj._matrix._coord_._col.ValI, string> GetDict(
			WorkbookPart workbookPart,
			Worksheet sheetData
			,
			nilnul.obj.matrix.block.set_.sed.Bounding _belt
			
		) {

			var r = new Dictionary<nilnul.obj._matrix._coord_._col.ValI, string>(
				nilnul.obj._matrix._coord_._col.val.Eq.Singleton
			);

			var firstCol = _belt.bounding.colRange.lower;

			var txts = GetTxts(workbookPart, sheetData, _belt).ToArray();

			for (int i = 0; i < txts.Length; i++)
			{

				r.Add(

					nilnul.obj._matrix._coord_._col.val.convert_._ShiftX.Shift(firstCol
					,
					i
					)
					,
					txts[i]

				);

			}
			return r;
		}

		static public IEnumerable< string > GetTxts(
			WorkbookPart workbookPart,
			Worksheet sheetData
			,
			nilnul.obj.matrix.block.set_.sed.Bounding _belt
		) {

			return GetTxts(workbookPart, sheet.DweltX.GetSheetData(sheetData), _belt);
				
				
		}

		[Obsolete(nameof(col._TxtX))]

		//todo to retor after being refactored to use Sheet data
		public static string GetTxt(
			WorkbookPart workbookPart,
			Worksheet worksheet, 
			nilnul.obj.matrix.block.Set minBelt
			,
			nilnul.obj._matrix._coord_._col.ValI col
		)
		{
			var intersectedBlocks = minBelt.Where(b=>b.colRange.contain(col)).OrderBy(
				b=>b
				,
				nilnul.obj.matrix.block.comp_.ColLower.Singleton
			 
			);

			return nilnul.txt.accumulate_.Concat.Singleton.accumulate(

				intersectedBlocks.Select(
					b=>
				
					nilnul.fs.excel.doc.sheet.dwelt.closure._TxtX.Get( 
						workbookPart, 
						doc.sheet.DweltX.GetSheetData(worksheet) ,
						b
					)
				)
			);

			//throw new NotImplementedException();
		}

		[Obsolete(nameof(col._TxtX))]
		public  static string GetTxt(
			WorkbookPart workbookPart,
			SheetData worksheet, 
			nilnul.obj.matrix.block.Set minBelt
			,
			nilnul.obj._matrix._coord_._col.ValI col
		)
		{
			var intersectedBlocks = minBelt.Where(b=>b.colRange.contain(col)).OrderBy(
				b=>b
				,
				nilnul.obj.matrix.block.comp_.ColLower.Singleton
			 
			);

			return nilnul.txt.accumulate_.Concat.Singleton.accumulate(

				intersectedBlocks.Select(
					b=>
				
					nilnul.fs.excel.doc.sheet.dwelt.closure._TxtX.Get(  
						workbookPart,
						worksheet ,
						b
					)
				)
			);

			//throw new NotImplementedException();
		}


	}
}
