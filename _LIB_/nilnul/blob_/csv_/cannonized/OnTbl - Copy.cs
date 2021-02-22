using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System;
using System.Data;
using System.IO;
using System.Text;

namespace nilnul.blob_.csv_.cannonized
{
	static public class _OnTbl_semicolonEndX
	{


		/// <summary>
		/// 
		/// </summary>
		/// <param name="filePath">
		///	using ";"
		///	"; " is at the end
		/// </param>
		/// <returns></returns>
		public static DataTable ImportFromCsvFile(string filePath)
		{
			var rowIndex = 0;
			var data = new DataTable();

			try
			{
				using (var reader = new StreamReader(File.OpenRead(filePath)))
				{
					while (!reader.EndOfStream)
					{
						var line = reader.ReadLine();
						var values = line.Substring(0, line.Length - 1).Split(';');

						foreach (var item in values)
						{
							if (string.IsNullOrEmpty(item) || string.IsNullOrWhiteSpace(item))
							{
								throw new Exception("Value can't be empty");
							}

							if (rowIndex == 0)
							{
								data.Columns.Add(item);
							}
						}

						if (rowIndex > 0)
						{
							if (values.Length != data.Columns.Count)
							{
								throw new Exception("Row is shorter or longer than title row");
							}
							data.Rows.Add(values);
						}

						rowIndex++;

					}
				}

				//var differentValuesOfLastColumn = MyAttribute.GetDifferentAttributeNamesOfColumn(data, data.Columns.Count - 1);

				//if (differentValuesOfLastColumn.Count > 2)
				//{
				//	throw new Exception("The last column is the result column and can contain only 2 different values");
				//}
			}
			catch (Exception ex)
			{
				throw;
				//DisplayErrorMessage(ex.Message);
				//data = null;
			}

			return data;
			// if no rows are entered or data == null, return null
			//return data?.Rows.Count > 0 ? data : null;
		}
		public static DataTable ImportFromBlob(string filePath)
		{
			var lines = nilnul.txts.fro_.txt_.split_._LineX.Line_removeWhite(filePath);

			var rowIndex = 0;
			var data = new DataTable();



			try
			{
				using (var reader = lines.GetEnumerator())
				{
					while (reader.MoveNext())
					{
						var line = reader.Current;
						var values = line.Substring(0, line.Length - 1).Split(';');

						foreach (var item in values)
						{
							if (string.IsNullOrEmpty(item) || string.IsNullOrWhiteSpace(item))
							{
								throw new Exception("Value can't be empty");
							}

							if (rowIndex == 0)
							{
								data.Columns.Add(item);
							}
						}

						if (rowIndex > 0)
						{
							if (values.Length != data.Columns.Count)
							{
								throw new Exception("Row is shorter or longer than title row");
							}
							data.Rows.Add(values);
						}

						rowIndex++;

					}
				}

				//var differentValuesOfLastColumn = MyAttribute.GetDifferentAttributeNamesOfColumn(data, data.Columns.Count - 1);

				//if (differentValuesOfLastColumn.Count > 2)
				//{
				//	throw new Exception("The last column is the result column and can contain only 2 different values");
				//}
			}
			catch (Exception ex)
			{
				throw;
				//DisplayErrorMessage(ex.Message);
				//data = null;
			}

			return data;
			// if no rows are entered or data == null, return null
			//return data?.Rows.Count > 0 ? data : null;
		}

		public static void ExportToCsvFile(DataTable data, string filePath)
		{
			if (data.Columns.Count == 0)
			{
				throw new Exception("Nothing to export");
			}

			var sb = new StringBuilder();

			// add titles to the string builder
			foreach (var item in data.Columns)
			{
				// seperate values with a ;
				sb.AppendFormat($"{item};");
			}

			sb.AppendLine();

			// add every row to the string builder
			for (var i = 0; i < data.Rows.Count; i++)
			{
				for (var j = 0; j < data.Columns.Count; j++)
				{
					// seperate values with a ;
					sb.AppendFormat($"{data.Rows[i][j]};");
				}

				sb.AppendLine();
			}

			File.WriteAllText(filePath, sb.ToString());

			//Console.ForegroundColor = ConsoleColor.Green;
			//Console.WriteLine("Data sucessfully exported");
			//Console.ResetColor();
		}


	}
}