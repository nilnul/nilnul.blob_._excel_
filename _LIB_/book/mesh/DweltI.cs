using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel.doc.sheet
{
	/// <summary>
	/// Bloc0nul.
	/// some cells in excel are never populated, hence these cells will not show in blob.
	/// for all populated cells, if any, there is a minimu block containing it. This block is called "Dwelt". but dwelt can be empty, in this case there is no block.
	/// </summary>
	public interface IDwelt
	{
	}
}
