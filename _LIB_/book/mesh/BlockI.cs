using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet
{
	/// <summary>
	/// correspons to "sheetData". 
	/// a block of which the direction of drag doesn't matter.
	/// the dwelled meaningful part. the index of this part is reintiated; that is , the cell [0,0] is the uppler left corner. In comparison with the sheet, the index has to be shifted/translated.
	/// 
	/// </summary>
	public interface BlockI
		:nilnul.obj.matrix.BlockI
	{


	}


}
