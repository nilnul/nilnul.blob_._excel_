namespace nilnul.blob_._excel_._TEST_.doc.sheet_.first.bloc.toTbl
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("nilnul.Task1")]
    public partial class Task1
    {
        public long id { get; set; }

        public string note { get; set; }

        [Column("_time")]
        public DateTime? C_time { get; set; }

        [Column("_note")]
        public string C_note { get; set; }
    }
}
