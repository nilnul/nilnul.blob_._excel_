using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet.dwelt.row.cel._key._ret
{
	[Obsolete()]
	public class Row
	{

		private uint _index;

		public uint index
		{
			get { return _index; }
			set { _index = value; }
		}
		private Row(uint index)
		{
			_index = index;
		}

		private Row(int row):this((uint)(row))
		{
		}

		private Row(BigInteger bigInteger):this((uint)bigInteger)
		{
		}


		/// <summary>
		/// to one based
		/// </summary>
		/// <returns></returns>
		public override string ToString()
		{
			return (_index+1).ToString();
			;
		}

		/// <summary>
		/// name is one based
		/// </summary>
		/// <param name="name"></param>
		/// <returns></returns>
		static public Row Parse(string name) {

			return new _ret.Row( uint.Parse(name) -1);
		}

		static public  Row CreateFroRow(DocumentFormat.OpenXml.Spreadsheet.Row row) {
			return new Row(row.RowIndex - 1);
		}

		static public  Row CreateFroRow(int row) {
			return new Row(row - 1);
		}

		public static Row CreateFroRow(BigInteger val)
		{
			return new Row(
				val-1
			);
			throw new NotImplementedException();
		}
	}
}
