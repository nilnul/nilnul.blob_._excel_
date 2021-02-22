using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.fs.excel.doc.sheet.dwelt.closures_;
using DocumentFormat.OpenXml.Packaging;
using nilnul.fs.excel.doc.sheet.dwelt.closures;

namespace nilnul.blob_.excel.doc.sheet.dwelt.belt.cols.sel_
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
				c => col._ValX. GetTxt(
					workbookPart,
					sheetData
					,_belt.toOriginal()
					,c
					
				)
			);
		}

		static public IEnumerable< string > GetTxts(
			

			Belt _belt
		)
		{

			return _belt.cols.Select(
				c => col._ValX. GetTxt(
					
					_belt
					,c
					
				)
			);
		}

		

		static public IEnumerable< string > GetTxts(
			WorkbookPart workbookPart,
			Worksheet sheetData
			,
			nilnul.obj.matrix.block.set_.sed.Bounding _belt
		) {

			return GetTxts(workbookPart, nilnul.fs.excel.doc.sheet.DweltX.GetSheetData(sheetData), _belt);
				
				
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



		internal static Dictionary<nilnul.obj._matrix._coord_._col.ValI, string> GetDict(
			Belt belt
		)
		{
			

			var r = new Dictionary<nilnul.obj._matrix._coord_._col.ValI, string>(
				nilnul.obj._matrix._coord_._col.val.Eq.Singleton
			);

			var firstCol = belt.cols.lower;

			var txts = GetTxts( belt).ToArray();

			for (int i = 0; i < txts.Length; i++)
			{

				r.Add(

					nilnul.obj._matrix._coord_._col.val.convert_._ShiftX.Shift(
						new nilnul.obj._matrix._coord_._col.Val(
						
						firstCol
						)
					,
					i
					)
					,
					txts[i]

				);

			}
			return r;
		}
	}
}
