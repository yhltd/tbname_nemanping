
namespace model
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;
    //[Table("HJ_CMS_Cart")]
    public partial class UserTable
    {
        [Key]
        public string id { get; set; }
        public string name { get; set; }
        public string pwd { get; set; }
        public bool Loginpd { get; set; }
    }
}
