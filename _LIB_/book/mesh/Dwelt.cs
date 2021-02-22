using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nilnul.obj._matrix.coord.duo.be_.le.vow;
using nilnul.obj.matrix;

namespace nilnul.blob_.excel.book.mesh
{
	/// <summary>
	/// the sheetData. a block
	/// </summary>
	public class Dwelt:
		nilnul.obj.Box1<excel.book.Mesh>
		
	{
		private excel.book.Mesh	 _mesh;

		public excel.book.Mesh mesh
		{
			get { return _mesh; }
			set { _mesh = value; }
		}

		public nilnul.obj.mesh.Bloc bloc {
			get {
				return nilnul.blob_.excel.book.mesh.dwelt._DiagX.Bloc(_mesh.worksheetPart.Worksheet);

				//return nilnul.fs.excel.doc.sheet.dwelt._AsDiagX.Get(_mesh.worksheetPart.Worksheet);
			}
		}
		



		public Dwelt(Mesh val) : base(val)
		{
		}

		 public Dwelt(SpreadsheetDocument doc, Worksheet _worksheet__mustBeInDoc)
			:this(new Mesh(doc,_worksheet__mustBeInDoc))

		{
			
		}



	}
}
