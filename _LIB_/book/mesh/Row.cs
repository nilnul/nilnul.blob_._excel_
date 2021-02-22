using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.book.mesh
{
	public class Major
	{
		public book.Mesh _mesh;
		public book.Mesh mesh {
			get { return _mesh; }
		}

		private Row _row;

		public Row row
		{
			get { return _row; }
			set { _row = value; }
		}


		public Major(book.Mesh mesh, Row row)
		{
			_mesh = mesh;
			_row = row;
		}


	}
}
