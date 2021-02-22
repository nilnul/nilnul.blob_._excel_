using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet.dwelt._cel
{
	/// <summary>
	/// when DataType is null, it is date or number.
	/// you can see this formatId 14 translates to short date. For a complete list of formats please refer to ECMA-376 documentation for Office Open XML formats (table should be buried somewhere inside part 4). 
	/// </summary>
	public enum NumDateFormatId
	{
		General = 0,
		Number = 1,
		Decimal = 2,
		Currency = 164,
		Accounting = 44,

		DateShort = 14,
		DateLong = 165,
		Time = 166,

		Percentage = 10,
		Fraction = 12,
		Scientific = 11,

		Text = 49
	}


}
