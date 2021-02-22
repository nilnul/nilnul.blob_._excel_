using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nilnul.blob_.excel
{
	/// <summary>
	/// an excel doc
	/// </summary>
	static public class _BookX
	{
		static public WorkbookPart GetWorkbookPart(SpreadsheetDocument doc) {
			return doc.WorkbookPart;
		}

		static public WorkbookPart GetWorkbookPart(string doc, bool isEditable=false) {
			return GetWorkbookPart( SpreadsheetDocument.Open(doc, isEditable));
		}


		static public SpreadsheetDocument Open(string doc, bool isEditable=true) {
			return  SpreadsheetDocument.Open(doc, isEditable);
		}



	}

	public class Book : nilnul.obj.Box1<SpreadsheetDocument>
	{
		public Book(SpreadsheetDocument val) : base(val)
		{
		}

		public SpreadsheetDocument blob {
			get { return boxed; }
		}


	}
}
