namespace nilnul.blob_._excel_._TEST_.doc.sheet_.first.bloc.toTbl
{
	using System;
	using System.Data.Entity;
	using System.ComponentModel.DataAnnotations.Schema;
	using System.Linq;

	public partial class Model1 : DbContext
	{
		public Model1()
			: base("name=Model1")
		{
		}
		public Model1(string s ):base(s)
		{

		}

		public virtual DbSet<Task1> Task1 { get; set; }

		protected override void OnModelCreating(DbModelBuilder modelBuilder)
		{
		}
	}
}
