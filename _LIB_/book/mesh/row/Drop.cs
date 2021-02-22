using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.book.mesh.row
{
	static public class _DropX
	{
		static public void Drop(excel.book.Mesh mesh, DocumentFormat.OpenXml.Spreadsheet.Row row)
		{

			row.Remove(); //all cells under are also removed.

			var rows = doc.sheet.dwelt.rows_.Populated.Get(mesh.worksheetPart.Worksheet);

			var mergeCells = excel.book.mesh._CelMergedS.Collection(mesh);

			//get the rows of the mergeCels

			var rowOrd = new nilnul.num.ord_.OneBased(
							row.RowIndex.Value
						);

			foreach (var item in _CelMergedS.Enumerate(mergeCells).ToArray())
			{
				var megeDiag = book.mesh._CelMergedX.ToDiag(item);

				if (
					nilnul.num.ord.comparer.Re1.Singleton.lt(

						megeDiag.end.row,
						rowOrd

					)
				)
				{
					continue;


				}

				if (
					nilnul.num.ord.comparer.Re1.Singleton.gt(

						megeDiag.begin.row
						,
						rowOrd

					)



				)
				{
					item.Reference = new excel.book.mesh._bloc.Diag(
						  _cel.Coord.OvColRowNilBased(megeDiag.begin.col, megeDiag.begin.row.toNum().toBigint().en - 1)
						,
						  _cel.Coord.OvColRowNilBased(megeDiag.end.col, megeDiag.end.row.toNum().toBigint().en - 1)


					).ToString();

					continue;

				}


				var delLine = obj.mesh.bloc_.diag.nulable.op_.unary_._DelLineX.Diag_ofLineOneBased(
					megeDiag
					,

					row.RowIndex.Value
				);

				if (delLine is null)
				{
					item.Remove();
				}
				else
				{
					//if (!nilnul.obj.mesh.bloc.Eq.Singleton.Equals( delLine, megeDiag  ))
					//{
					item.Reference = new excel.book.mesh._bloc.Diag(delLine).ToString();

					//}

				}



			}

			/// move up lines
			/// 
			foreach (var r in rows)
			{
				if (r.RowIndex.HasValue)
				{
					if (r.RowIndex.Value > row.RowIndex.Value)
					{
						//move up merege cells in this row



						IEnumerable<Cell> cells = r.Elements<Cell>().ToList();
						if (cells != null)
						{
							foreach (Cell cell in cells)
							{
								var cr = cell.CellReference.Value;

								int newIndexRow = Convert.ToInt32(Regex.Replace(cr, @"[^\d]+", "")) - 1;

								var rRemoveNum = Regex.Replace(cr, @"[\d-]", "");

								cell.CellReference.Value = $"{rRemoveNum}{newIndexRow}";
							}
						}
						r.RowIndex.Value -= 1;

					}

				}
			}
			//mesh.worksheetPart.Worksheet.Save();
			//mesh.book.Close();

		}

		static public void Drop_ofRowOneBased(excel.book.Mesh mesh, int row)
		{
			var r = book.mesh.rows_.dwelt._ChooseX.Row0nul_ofRowOneBased(mesh, row);

			Drop(mesh, r);

		}

	}
}
