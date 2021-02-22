using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.fs.excel.doc.sheet.read_
{
	class Class1
	{

		[Obsolete("fro web. revised such that it's not working anymore",true)]
		public static DataTable ExcelWorksheetToDataTable(string pathFilename, string worksheetName)
		{
			DataTable dt = new DataTable(worksheetName);

			using (SpreadsheetDocument document = SpreadsheetDocument.Open(pathFilename, false))
			{
				// Find the sheet with the supplied name, and then use that 
				// Sheet object to retrieve a reference to the first worksheet.
				Sheet theSheet = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName).FirstOrDefault();
				if (theSheet == null)
					throw new Exception("Couldn't find the worksheet: " + worksheetName);

				// Retrieve a reference to the worksheet part.
				WorksheetPart wsPart = (WorksheetPart)(document.WorkbookPart.GetPartById(theSheet.Id));
				Worksheet workSheet = wsPart.Worksheet;

				string dimensions = workSheet.SheetDimension.Reference.InnerText;       //  Get the dimensions of this worksheet, eg "B2:F4"

				int numOfColumns = 0;
				int numOfRows = 0;
				_DimX.CalculateTillLowRight(dimensions, ref numOfColumns, ref numOfRows);

				System.Diagnostics.Trace.WriteLine(string.Format("The worksheet \"{0}\" has dimensions \"{1}\", so we need a DataTable of size {2}x{3}.", worksheetName, dimensions, numOfColumns, numOfRows));

				SheetData sheetData = workSheet.GetFirstChild<SheetData>();
				IEnumerable<Row> rows = sheetData.Descendants<Row>();

				string[,] cellValues = new string[numOfColumns, numOfRows];

				int colInx = 0;
				int rowInx = 0;
				string value = "";
				SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;

				//  Iterate through each row of OpenXML data
				foreach (Row row in rows)
				{
					for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
					{
						//  *DON'T* assume there's going to be one XML element for each item in each row...
						Cell cell = row.Descendants<Cell>().ElementAt(i);
						if (cell.CellValue == null || cell.CellReference == null)
							continue;                       //  eg when an Excel cell contains a blank string

						//  Convert this Excel cell's CellAddress into a 0-based offset into our array (eg "G13" -> [6, 12])
						colInx = col._ColX.GetColumnIndexByName(cell.CellReference);             //  eg "C" -> 2  (0-based)
						rowInx = sheet.row.cel.GetIndex.GetRowIndexFromCellAddress(cell.CellReference) - 1;     //  Needs to be 0-based  

						//  Fetch the value in this cell
						value = cell.CellValue.InnerXml;
						if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
						{
							value = stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
						}

						cellValues[colInx, rowInx] = value;
					}
					//dt.Rows.Add(dataRow);
				}

				//  Copy the array of strings into a DataTable
				for (int col = 0; col < numOfColumns; col++)
					dt.Columns.Add("Column_" + col.ToString());

				for (int row = 0; row < numOfRows; row++)
				{
					DataRow dataRow = dt.NewRow();
					for (int col = 0; col < numOfColumns; col++)
					{
						dataRow.SetField(col, cellValues[col, row]);
					}
					dt.Rows.Add(dataRow);
				}

//#if DEBUG
//				//  Write out the contents of our DataTable to the Output window (for debugging)
//				string str = "";
//				for (rowInx = 0; rowInx < maxNumOfRows; rowInx++)
//				{
//					for (colInx = 0; colInx < maxNumOfColumns; colInx++)
//					{
//						object val = dt.Rows[rowInx].ItemArray[colInx];
//						str += (val == null) ? "" : val.ToString();
//						str += "\t";
//					}
//					str += "\n";
//				}
//				System.Diagnostics.Trace.WriteLine(str);
//#endif
				return dt;
			}
		}


	}
}
