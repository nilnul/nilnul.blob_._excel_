using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet.dwelt.closures
{
	/// <summary>
	/// divide the closures into belts
	/// </summary>
	/// 
	[Obsolete(nameof(nilnul.blob_.excel.doc.sheet.dwelt.Belts))]
	public class Belts
	{



		private Worksheet _worksheet;

		public Worksheet worksheet
		{
			get { return _worksheet; }
			set { _worksheet = value; }
		}

		public IEnumerable<nilnul.obj.matrix.block.set_.sed.Bounding> enumerate()
		{

			return Enumerate(_worksheet);
		}
		



		static public IEnumerable<_sheet._coord_._row.val.Range> HorizontalBoundS(Worksheet worksheet)
		{
			var enumerator =  dwelt._rows._GetX.RowIndexS(worksheet).GetEnumerator();

			while (enumerator.MoveNext())
			{
				var belt =row_.mergedNoPre.belt.range.bound_._VerticalX._Get(
					worksheet,
					enumerator.Current
				);

				yield return belt;

				nilnul.obj.seq.enumerator._MoveNextX.MoveNext(
					enumerator,
					(int)belt.count.toBigint() - 1 //we 'll move one row more later.
				);
			} 
		}

		static public IEnumerable<Belt> EnumerateInBelt(WorkbookPart workbookPart,Worksheet worksheet)
		{

			return HorizontalBoundS(worksheet).Select(
				x=> new Belt(new nilnul.fs.excel.doc.sheet.Dwelt(workbookPart,worksheet), x )
			);
		}




		static public IEnumerable<nilnul.obj.matrix.block.set_.sed.Bounding> Enumerate(Worksheet worksheet)
		{
			var enumerator =  dwelt._rows._GetX.RowIndexS(worksheet).GetEnumerator();

			while (enumerator.MoveNext())
			{
				var belt = new nilnul.obj.matrix.block.set_.sed.Bounding(
					new nilnul.obj.matrix.block.set_.Sed(
						row_.mergedNoPre._BeltX.GetSet(
							worksheet,
							enumerator.Current
						)
					)
				);

				yield return belt;

				nilnul.obj.seq.enumerator._MoveNextX.MoveNext(
					enumerator,
					(int)belt.bounding.rowRange.count.toBigint() - 1 //we 'll move one row more later.
				);
			} 
		}

	}
}
