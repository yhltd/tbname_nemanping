//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using System.Configuration;
//using model;
//using System.Data.SqlClient;
//namespace clsBuiness
//{
//    using System;
//    using System.Data.Entity;
//    using System.ComponentModel.DataAnnotations.Schema;
//    using System.Linq;
//    public partial class clsAllnew : DbContext
//    {
//        public clsAllnew()
//            : base("name=Purchasing")
//        {
        
//        }
//        public DbSet<UserTable> UserTable { get; set; }
//        protected override void OnModelCreating(DbModelBuilder modelBuilder)
//        {
//            modelBuilder.Entity<UserTable>()
//                 .Property(u=>u.Id)
//                 .IsUnicode(false);
//        }
        ////protected  sqlhelper sqldb = new sqlhelper();
        //protected SqlDataReader rd;
        //public clsAllnew()
        //{
        //    sqlhelper.OpenConn();
        //}
        ////查询UserTable表
        //public List<UserTable> selectUserTable()
        //{
        //    List<UserTable> list = new List<UserTable>();
        //    string sql = "select * from userTable";
        //    rd = sqlhelper.GetDataReaderValue(sql);
        //    while (rd.Read()) 
        //    {
        //        UserTable ut = new UserTable();
        //        ut.id =rd.GetInt32(rd.GetOrdinal("Id")).ToString();
//        //        ut.name = rd.GetString(rd.GetOrdinal("name"));
//        //        ut.pwd = rd.GetString(rd.GetOrdinal("pwd"));
//        //        list.Add(ut);
//        //    }
//        //    sqlhelper.CloseConn();
//        //    return list;
//        //}
//    }
//}
