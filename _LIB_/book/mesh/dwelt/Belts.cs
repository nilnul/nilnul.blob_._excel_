using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;

namespace nilnul.blob_.excel.doc.sheet.dwelt
{
	/// <summary>
	/// divide the closures into belts
	/// </summary>
	public class Belts
	{



		private Worksheet _worksheet;

		public Worksheet worksheet
		{
			get => _worksheet;
			set => _worksheet = value;
		}

		public IEnumerable<nilnul.obj.matrix.block.set_.sed.Bounding> enumerate()
		{

			return Enumerate(_worksheet);
		}




		public static IEnumerable<nilnul.fs.excel.doc._sheet._coord_._row.val.Range> HorizontalBoundS(Worksheet worksheet)
		{
			IEnumerator<obj._matrix._coord_._row.Val> enumerator = nilnul.fs.excel.doc.sheet.dwelt._rows._GetX.RowIndexS(worksheet).GetEnumerator();

			while (enumerator.MoveNext())
			{
				fs.excel.doc._sheet._coord_._row.val.Range belt = nilnul.fs.excel.doc.sheet.dwelt.row_.mergedNoPre.belt.range.bound_._VerticalX._Get(
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

		public static IEnumerable<Belt> EnumerateInBelt(WorkbookPart workbookPart, Worksheet worksheet)
		{

			return HorizontalBoundS(worksheet).Select(
				x => new Belt(new nilnul.blob_.excel.doc.sheet.Dwelt(workbookPart, worksheet),
				 new num.ord_.oneBased.Bound(
					  nilnul.num.ord_.OneBased.FroNilBased(
						new Num(
							x.lower.toNum()
							)
						)

					,
					  nilnul.num.ord_.OneBased.FroNilBased(
						new Num(
							x.lower.toNum()
							)
						)


				)
				)
			);
		}




		public static IEnumerable<nilnul.obj.matrix.block.set_.sed.Bounding> Enumerate(Worksheet worksheet)
		{
			IEnumerator<obj._matrix._coord_._row.Val> enumerator = nilnul.fs.excel.doc.sheet.dwelt._rows._GetX.RowIndexS(worksheet).GetEnumerator();

			while (enumerator.MoveNext())
			{
				obj.matrix.block.set_.sed.Bounding belt = new nilnul.obj.matrix.block.set_.sed.Bounding(
					new nilnul.obj.matrix.block.set_.Sed(
						nilnul.fs.excel.doc.sheet.dwelt.row_.mergedNoPre._BeltX.GetSet(
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
