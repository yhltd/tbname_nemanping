﻿//------------------------------------------------------------------------------
// <auto-generated>
//    此代码是根据模板生成的。
//
//    手动更改此文件可能会导致应用程序中发生异常行为。
//    如果重新生成代码，则将覆盖对此文件的手动更改。
// </auto-generated>
//------------------------------------------------------------------------------

namespace clsBuiness
{
    using System;
    using System.Data.Entity;
    using System.Data.Entity.Infrastructure;
    
    public partial class nemanpingEntities3 : DbContext
    {
        public nemanpingEntities3()
            : base("name=nemanpingEntities3")
        {
        }
    
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            throw new UnintentionalCodeFirstException();
        }
    
        public DbSet<C_PANT> C_PANT { get; set; }
        public DbSet<CaiDan> CaiDan { get; set; }
        public DbSet<CaiDan_C_PANT> CaiDan_C_PANT { get; set; }
        public DbSet<CaiDan_D_PANT> CaiDan_D_PANT { get; set; }
        public DbSet<CaiDan_RGL2> CaiDan_RGL2 { get; set; }
        public DbSet<CaiDan_RGLJ> CaiDan_RGLJ { get; set; }
        public DbSet<CaiDan_SLIM> CaiDan_SLIM { get; set; }
        public DbSet<ChiMa_Dapeibiao> ChiMa_Dapeibiao { get; set; }
        public DbSet<D_PANT> D_PANT { get; set; }
        public DbSet<DanHao> DanHao { get; set; }
        public DbSet<GongHuoFang> GongHuoFang { get; set; }
        public DbSet<JiaGongChang> JiaGongChang { get; set; }
        public DbSet<KuanShiBiao> KuanShiBiao { get; set; }
        public DbSet<KuCun> KuCun { get; set; }
        public DbSet<MianFuLiaoDingGouDan> MianFuLiaoDingGouDan { get; set; }
        public DbSet<PeiSe> PeiSe { get; set; }
        public DbSet<RGL2> RGL2 { get; set; }
        public DbSet<RGLJ> RGLJ { get; set; }
        public DbSet<Sehao> Sehao { get; set; }
        public DbSet<SLIM> SLIM { get; set; }
        public DbSet<UserTable> UserTable { get; set; }
    }
}
