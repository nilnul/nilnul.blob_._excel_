using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using nilnul.obj._matrix._coord_._col;
using nilnul.obj._matrix._coord_._row;

namespace nilnul.fs.excel.doc._sheet._coord_._row
{
	/// <summary>
	/// y or vertical axis
	/// </summary>
	public class Val
		: nilnul.obj._matrix._coord_._row.val_.Positive

	{

		public Val(nilnul.obj._matrix._coord_._row.ValI row):base(row)
		{
		}

		public Val(UInt32Value rowIndex):base((uint)rowIndex)
		{
		}

		public Val(uint index):base(index)
		{

		}

		public Val(NumI num):base(
			num
			
		)
		{
		}

		public Val(BigInteger x) : base(x)
		{
		}



		//public Row(int index):base(new num_.Positive( index))
		//{
		//}

		public Val(DocumentFormat.OpenXml.Spreadsheet.Row row)
			:base( (uint)  row.RowIndex)
		{

		}

		
		static public Val CreateFro0Based(int x) {
			return new Val(x + 1);
		}

		internal static Val CreateFro0Based(NumI maxExt)
		{
			return new Val(
				nilnul.num.convert_.Inc.Singleton.convert(maxExt)
			);
			throw new NotImplementedException();
		}
	}
}
