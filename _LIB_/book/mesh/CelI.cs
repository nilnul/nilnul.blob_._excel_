using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.book.mesh
{
	public interface CelI 
		
	{
	}

	public class Cel
	{
		public book.Mesh _mesh;
		public book.Mesh mesh
		{
			get { return _mesh; }
		}

		private Cell _cell;

		public Cell cell
		{
			get { return _cell; }
			set { _cell = value; }
		}


		public Cel(book.Mesh mesh, Cell row)
		{
			_mesh = mesh;
			_cell = row;
		}


	}
}
