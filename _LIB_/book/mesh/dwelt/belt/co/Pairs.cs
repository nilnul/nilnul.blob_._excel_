using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using nilnul.fs.excel.doc._sheet._coord_._row;
using DocumentFormat.OpenXml.Spreadsheet;

namespace nilnul.blob_.excel.doc.sheet.dwelt.belt.co
{
	/// <summary>
	/// filter the first belt with the second
	/// </summary>
	static public class _PairsX
	{
		/// <summary>
		/// get the relation between src and target by enumerating cols.
		/// This is a total surjective relation. may be polygomy; it also can be polyandry as there may be white keys.
		/// </summary>
		/// <param name="doc"></param>
		/// <param name="worksheet"></param>
		/// <param name="row"></param>
		/// <returns></returns>
		public static IEnumerable<nilnul.txt.Duo> GetTxtDuoS(
			SpreadsheetDocument doc
			,
			Worksheet worksheet
			,
			nilnul.obj._matrix._coord_._row.ValI row
			,
			nilnul.obj._matrix._coord_._row.ValI rowFilter

		)
		{

			var dim = nilnul.blob_.excel.doc.sheet.dwelt._DiagX.Get(worksheet);

			var belts = nilnul.blob_.excel.doc.sheet.dwelt.Belts.Enumerate(
				worksheet
			);


			// get the belt from row

			//var belt = belts.FirstOrDefault(b => nilnul.obj._matrix._coord_._row.val.Eq.Singleton.Equals(
			//	b.bounding.rowRange.lower
			//	,
			//	row
			//));
			var belt = belts.FirstOrDefault(
				b => b.bounding.rowRange.contain(
					row
				)
			);

			//get the next belt, which is the target
			var targets = belts.FirstOrDefault(
				b => b.bounding.rowRange.contain(
					rowFilter
				)
			);

			for (
				nilnul.obj._matrix._coord_._col.ValI i = dim.colRange.lower;
				nilnul.obj._matrix._coord_._col.val.comp.Re.Singleton.le(i, dim.colRange.upper);
				i = nilnul.obj._matrix._coord_._col.val.convert_.Inc.Singleton.convert(i)
			)
			{
				yield return new txt.Duo(

					nilnul.txt.op_.PurgeWhite.Singleton.op(
						nilnul.fs.excel.doc.sheet.dwelt.closures.belt.col._TxtX.GetTxt(
							doc.WorkbookPart
							,
							worksheet,
							belt,
							i
						)
					)
					,
					nilnul.fs.excel.doc.sheet.dwelt.closures.belt.col._TxtX.GetTxt(
						doc.WorkbookPart
						,
						worksheet,
						targets,
						i
					)
				);
			}
		}
	
	}
}