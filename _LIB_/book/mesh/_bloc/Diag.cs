using nilnul.num.ord;
using nilnul.num.ord_;
using nilnul.num.ord_.oneBased_.bijective_;
using nilnul.obj.mesh._cel;
using nilnul.obj.mesh._cel.coord.co.be_.le.vow;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.book.mesh._bloc
{
	public class Diag :
		nilnul.obj.mesh.Bloc
	{
		

		public Diag(_cel.Coord position1, _cel.Coord position2) : base(position1, position2)
		{
		}

		

		public Diag(Ee val) : this(
			new _cel.Coord(val.ee.component)
			,
			new _cel.Coord(val.ee.component1)



		)
		{
		}
	}
}
