using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.book.mesh.dwelt
{
	/*Horizon sounds like chinese stroke "Heng"*/
	public class Horizon : nilnul.obj.Box_pub<Dwelt>
	{
		private Row _row;

		public Row row
		{
			get { return _row; }
			set { _row = value; }
		}

		public Horizon(Dwelt val, Row _mustBeInDwelt) : base(val)
		{
			_row = _mustBeInDwelt;
		}
	}
}
