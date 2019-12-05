using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using clsBuiness;
using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Text.RegularExpressions;
//using NPOI.HSSF.UserModel;
namespace logic
{
    public class clsAllnewLogic
    {
        public clsAllnewLogic()
        {
            
        }
        public clsAllnewLogic(string a , string b) 
        {
            Console.WriteLine(a);
            Console.WriteLine(b);
        }


        public bool IsNumber(String strNumber)
        {
            Regex objNotNumberPattern=new Regex("[^0-9.-]");
            Regex objTwoDotPattern=new Regex("[0-9]*[.][0-9]*[.][0-9]*");
            Regex objTwoMinusPattern=new Regex("[0-9]*[-][0-9]*[-][0-9]*");
            String strValidRealPattern="^([-]|[.]|[-.]|[0-9])[0-9]*[.]*[0-9]+$";
            String strValidIntegerPattern="^([-]|[0-9])[0-9]*$";
            Regex objNumberPattern =new Regex("(" + strValidRealPattern +")|(" + strValidIntegerPattern + ")");
            return !objNotNumberPattern.IsMatch(strNumber) &&
                   !objTwoDotPattern.IsMatch(strNumber) &&
                   !objTwoMinusPattern.IsMatch(strNumber) &&
                  objNumberPattern.IsMatch(strNumber);
        }
        public bool Login(string name , string pwd) 
        {
            using (nemanpingEntities3 can = new nemanpingEntities3())
            {
                
                var user = can.UserTable.First(u => u.Name.Equals(name) && u.Pwd.Equals(pwd));
                return true;
            }
            
        }
        #region 功能：色号录入

            #region 查询色号
            public List<Sehao> selectSehao()
            {
                using (nemanpingEntities3 can = new nemanpingEntities3())
                {
                    List<Sehao> sehao = new List<Sehao>();
                    var select = from s in can.Sehao select s;
                    sehao = select.ToList<Sehao>();
                
                    return sehao;
                }
            }
            #endregion

            #region 添加色号
            public void insertSehao(DataTable dt)
            {
                using (nemanpingEntities3 can = new nemanpingEntities3())
                {
                    foreach (DataRow dr in dt.Rows)
                    {

                        
                        if (!dr[2].ToString().Equals(string.Empty))
                        {
                            string id = dr[2].ToString();
                            var select = from sc in can.Sehao where sc.SeHao1.Equals(id) select sc;
                            var target = select.FirstOrDefault<Sehao>();
                            if (target == null)
                            {
                                Sehao insets = new Sehao()
                                {
                                    Name = dr[1].ToString(),
                                    SeHao1 = dr[2].ToString()
                                };
                                can.Sehao.Add(insets);
                            }
                            else 
                            {
                                target.Name = dr[1].ToString();
                                target.SeHao1 = dr[2].ToString();
                            }
                        }
                 
                    }
                    can.SaveChanges();

                }
            }
            #endregion

            #region 读取色号表
            public List<Sehao> readerSehaoExcel(string fileName)
            {
                List<Sehao> list = new List<Sehao>();
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
                {
                    WorkbookPart wbPart = document.WorkbookPart;
                    List<Sheet> sheets = wbPart.Workbook.Descendants<Sheet>().ToList();
                    var versionSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                    WorksheetPart worksheetPart = (WorksheetPart)wbPart.GetPartById(versionSheet.Id);
                    int rowindex = 0;
                    foreach (Row row in worksheetPart.Worksheet.Descendants<Row>())
                    {
                        if (rowindex < 1)
                        {
                            rowindex++;
                            continue;
                        }
                        Sehao s = new Sehao();
                        foreach (Cell cell in row)
                        {
                            string rev = cell.CellReference.Value;
                            if (rev.StartsWith("B"))
                            {
                                s.Name = GetCellValue(wbPart, cell);
                            }
                            if (rev.StartsWith("A"))
                            {
                                s.SeHao1 = GetCellValue(wbPart, cell);
                            }
                        }
                        list.Add(s);
                    }
                    return list;
                }
            }
            #endregion 

            #region 删除色号表
            public void deleteSehao(List<int> id) 
            {
                using(nemanpingEntities3 npe = new nemanpingEntities3())
                {
                    foreach (int strid in id) 
                    {
                        if (strid != null)
                        {
                            if (strid != 0) 
                            {
                                var select = from s in npe.Sehao where s.Id == strid select s;
                                foreach (var item in select)
                                {
                                    npe.Sehao.Remove(item);
                                }
                            }
                        }
                    }
                    npe.SaveChanges();
                }
                
            }
            #endregion
        #endregion

        #region 读取EXCEL表格相关代码

        public static string GetCellValue(WorkbookPart wbPart, Cell theCell)
    {
        string value = theCell.InnerText;
        //  String value1 = theCell.CellValue.InnerText;
        if (theCell.DataType != null)
        {
            switch (theCell.DataType.Value)
            {
                case CellValues.SharedString:
                    var stringTable = wbPart.
                        GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    if (stringTable != null)
                    {
                        value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                    }
                    break;

                case CellValues.Boolean:
                    switch (value)
                    {
                        case "0":
                            value = "FALSE";
                            break;
                        default:
                            value = "TRUE";
                            break;
                    }
                    break;
            }
        }
        return value;
    }
    public String GetValue(Cell cell, WorkbookPart wbPart)
    {
        SharedStringTablePart stringTablePart = wbPart.SharedStringTablePart;
        if (cell.ChildElements.Count == 0)
            return null;
        String value = cell.CellValue.InnerText;
        if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
            return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
        return value;
    } 
    #endregion
        public List<PeiSe> selectps ()
        {
            using (nemanpingEntities3 npe = new nemanpingEntities3())
            {
                //List<PeiSe> list = new List<PeiSe>();
                List<PeiSe> list = npe.Database.SqlQuery<PeiSe>("select P.Id,Fabrics,PingMing,P.HuoHao,P.GuiGe, [61601C1] AS 'C61601C1',[61602C1] AS 'C61602C1' , [61603C1] AS 'C61603C1', [61605C1] AS 'C61605C1',[61606C1] AS 'C61606C1',[61607C1] AS 'C61607C1',[61609C1] AS 'C61609C1',[61611C1] AS'C61611C1',[61618C1] AS 'C61618C1',[61624C1] AS 'C61624C1',[61627C1] AS 'C61627C1',[61631C1] AS 'C61631C1',[61632C1] AS 'C61632C1', [61633C1] AS 'C61633C1' ,[61634C1] AS 'C61634C1' ,MianLiaoYanSe,P.Date from PeiSe p inner join DanHao d on p.PingMing = d.Name ").ToList<PeiSe>();
                //list = quer .ToList< PeiSe >();
                return list;
            }
        }
        #region 功能:尺码搭配表录入
            #region 查询尺码搭配表
        public List<ChiMa_Dapeibiao> SelectChiMaDapei(string strwhere) 
        {
            using(nemanpingEntities3 npe = new nemanpingEntities3())
            {
                List<ChiMa_Dapeibiao> list = new List<ChiMa_Dapeibiao>();

                if (strwhere.Equals(string.Empty))
                {
                    
                    var select = from cm in npe.ChiMa_Dapeibiao
                                 select cm;
                    list = select.ToList<ChiMa_Dapeibiao>();
                    
                }
                else 
                {
                    var select = from cm in npe.ChiMa_Dapeibiao
                                 where cm.BiaoName.Equals(strwhere)
                                 select cm;
                    list = select.ToList<ChiMa_Dapeibiao>();
                }
                return list;
            }
            
        }
        public List<RGL2> SelectChiMaDapei2()
        {
            using (nemanpingEntities3 npe = new nemanpingEntities3())
            {
                List<RGL2> list = new List<RGL2>();
                    var select = from cm in npe.RGL2
                                 select cm;
                    list = select.ToList<RGL2>();
                return list;
            }

        }
        public List<SLIM> SelectChiMaDapei3()
        {
            using (nemanpingEntities3 npe = new nemanpingEntities3())
            {
                List<SLIM> list = new List<SLIM>();

                //if (strwhere.Equals(string.Empty))
                //{

                var select = from cm in npe.SLIM
                                 select cm;
                    list = select.ToList<SLIM>();

                //}
                //else
                //{
                //    var select = from cm in npe.ChiMa_Dapeibiao
                //                 where cm.BiaoName.Equals(strwhere)
                //                 select cm;
                //    list = select.ToList<ChiMa_Dapeibiao>();
                //}
                return list;
            }

        }
        public List<RGLJ> SelectChiMaDapei4()
        {
            using (nemanpingEntities3 npe = new nemanpingEntities3())
            {
                List<RGLJ> list = new List<RGLJ>();

                //if (strwhere.Equals(string.Empty))
                //{

                    var select = from cm in npe.RGLJ
                                 select cm;
                    list = select.ToList<RGLJ>();

                //}
                //else
                //{
                //    var select = from cm in npe.ChiMa_Dapeibiao
                //                 where cm.BiaoName.Equals(strwhere)
                //                 select cm;
                //    list = select.ToList<ChiMa_Dapeibiao>();
                //}
                return list;
            }

        }
        public List<D_PANT> SelectChiMaDapei5()
        {
            using (nemanpingEntities3 npe = new nemanpingEntities3())
            {
                List<D_PANT> list = new List<D_PANT>();

                //if (strwhere.Equals(string.Empty))
                //{

                var select = from cm in npe.D_PANT
                                 select cm;
                list = select.ToList<D_PANT>();

                //}
                //else
                //{
                //    var select = from cm in npe.ChiMa_Dapeibiao
                //                 where cm.BiaoName.Equals(strwhere)
                //                 select cm;
                //    list = select.ToList<ChiMa_Dapeibiao>();
                //}
                return list;
            }

        }
        public List<C_PANT> SelectChiMaDapei6()
        {
            using (nemanpingEntities3 npe = new nemanpingEntities3())
            {
                List<C_PANT> list = new List<C_PANT>();

                //if (strwhere.Equals(string.Empty))
                //{

                    var select = from cm in npe.C_PANT
                                 select cm;
                    list = select.ToList<C_PANT>();

                //}
                //else
                //{
                //    var select = from cm in npe.ChiMa_Dapeibiao
                //                 where cm.BiaoName.Equals(strwhere)
                //                 select cm;
                //    list = select.ToList<ChiMa_Dapeibiao>();
                //}
                return list;
            }

        }

            #endregion

            #region 添加修改 尺码搭配表
                public void InsertChima(DataTable dt,string biaogeName) 
                {
                    using (nemanpingEntities3 npe = new nemanpingEntities3()) 
                    {
                        List<clsBuiness.ChiMa_Dapeibiao> cmdp = SelectChiMaDapei("");
                        foreach (DataRow dr in dt.Rows) 
                        {
                            List<clsBuiness.ChiMa_Dapeibiao> selectcm = new List<ChiMa_Dapeibiao>();
                            if (dr[0] !=null && dr[0].ToString() != "")
                            {
                                selectcm = cmdp.FindAll(f => f.LOT__面料 == dr[0].ToString() );
                                //if (dr[0] is DBNull || Convert.ToInt32(dr[0]) == 0)
                                //{
                            }
                            if(selectcm.Count<=0){
                                ChiMa_Dapeibiao cd = new ChiMa_Dapeibiao()
                                {
                                    BiaoName = biaogeName,
                                    LOT__面料 = dr[0].ToString(),
                                    STYLE_款式 = dr[1].ToString(),
                                    ART_货号 = dr[2].ToString(),
                                    COLOR_颜色 = dr[3].ToString(),
                                    COLOR__颜色编号 = dr[4].ToString(),
                                    JACKET_上衣_PANT_裤子 = dr[5].ToString(),
                                    C34R_28 = dr[6].ToString(),
                                    C36R_30 = dr[7].ToString(),
                                    C38R_32 = dr[8].ToString(),
                                    C40R___34 = dr[9].ToString(),
                                    C42R_36 = dr[10].ToString(),
                                    C44R_38 = dr[11].ToString(),
                                    C46R_40 = dr[12].ToString(),
                                    C48R_42 = dr[13].ToString(),
                                    C50R_44 = dr[14].ToString(),
                                    C52R_46 = dr[15].ToString(),
                                    C54R_48 = dr[16].ToString(),
                                    C56R_50 = dr[17].ToString(),
                                    C58R_52 = dr[18].ToString(),
                                    C60R_54 = dr[19].ToString(),
                                    C62R_56 = dr[20].ToString(),
                                    C36L_30 = dr[21].ToString(),
                                    C38L_32 = dr[22].ToString(),
                                    C40L_34 = dr[23].ToString(),
                                    C42L_36 = dr[24].ToString(),
                                    C44L_38 = dr[25].ToString(),
                                    C46L_40 = dr[26].ToString(),
                                    C48L_42 = dr[27].ToString(),
                                    C50L_44 = dr[28].ToString(),
                                    C52L_46 = dr[29].ToString(),
                                    C54L_48 = dr[30].ToString(),
                                    C56L_50 = dr[31].ToString(),
                                    C58L_52 = dr[32].ToString(),
                                    C60L_54 = dr[33].ToString(),
                                    C62L_56 = dr[34].ToString(),
                                    C34S_28 = dr[35].ToString(),
                                    C36S_30 = dr[36].ToString(),
                                    C38S_32 = dr[37].ToString(),
                                    C40S_34 = dr[38].ToString(),
                                    C42S_36 = dr[39].ToString(),
                                    C44S_38 = dr[40].ToString(),
                                    C46S_40 = dr[41].ToString(),
                                    DingdanHeji = dr[42].ToString()
                                };
                                npe.ChiMa_Dapeibiao.Add(cd);
                            }
                            else 
                            {
                                string id = dr[0].ToString();
                                var select = from cd in npe.ChiMa_Dapeibiao
                                             where cd.LOT__面料 ==id
                                             select cd;
                                var target = select.FirstOrDefault<ChiMa_Dapeibiao>();
                                target.LOT__面料 = dr[0].ToString();
                                target.STYLE_款式 = dr[1].ToString();
                                target.ART_货号 = dr[2].ToString();
                                target.COLOR_颜色 = dr[3].ToString();
                                target.COLOR__颜色编号 = dr[4].ToString();
                                target.JACKET_上衣_PANT_裤子 = dr[5].ToString();
                                target.C34R_28 = dr[6].ToString();
                                target. C36R_30 = dr[7].ToString();
                                target.C38R_32 = dr[8].ToString();
                                target.C40R___34 = dr[9].ToString();
                                target.C42R_36 = dr[10].ToString();
                                target.C44R_38 = dr[11].ToString();
                                target.C46R_40 = dr[12].ToString();
                                target.C48R_42 = dr[13].ToString();
                                target.C50R_44 = dr[14].ToString();
                                target.C52R_46 = dr[15].ToString();
                                target.C54R_48 = dr[16].ToString();
                                target.C56R_50 = dr[17].ToString();
                                target.C58R_52 = dr[18].ToString();
                                target.C60R_54 = dr[19].ToString();
                                target.C62R_56 = dr[20].ToString();
                                target.C36L_30 = dr[21].ToString();
                                target.C38L_32 = dr[22].ToString();
                                target.C40L_34 = dr[23].ToString();
                                target.C42L_36 = dr[24].ToString();
                                target.C44L_38 = dr[25].ToString();
                                target.C46L_40 = dr[26].ToString();
                                target.C48L_42 = dr[27].ToString();
                                target.C50L_44 = dr[28].ToString();
                                target.C52L_46 = dr[29].ToString();
                                target.C54L_48 = dr[30].ToString();
                                target.C56L_50 = dr[31].ToString();
                                target.C58L_52 = dr[32].ToString();
                                target.C60L_54 = dr[33].ToString();
                                target.C62L_56 = dr[34].ToString();
                                target.C34S_28 = dr[35].ToString();
                                target.C36S_30 = dr[36].ToString();
                                target.C38S_32 = dr[37].ToString();
                                target.C40S_34 = dr[38].ToString();
                                target.C42S_36 = dr[39].ToString();
                                target.C44S_38 = dr[40].ToString();
                                target.C46S_40 = dr[41].ToString();
                                target.DingdanHeji = dr[42].ToString();  
                            }
                        }
                        npe.SaveChanges();
                    }
                }
                public void InsertChima2(DataTable dt, string biaogeName)
                {
                    using (nemanpingEntities3 npe = new nemanpingEntities3())
                    {
                        List<clsBuiness.RGL2> cmdp = SelectChiMaDapei2();
                        foreach (DataRow dr in dt.Rows)
                        {
                            List<clsBuiness.RGL2> selectcm = new List<RGL2>();
                            //if (dr[0].ToString() != "")
                            //{
                                selectcm = cmdp.FindAll(f => f.LOT_ == dr[0].ToString() );
                                //if (dr[0] is DBNull || Convert.ToInt32(dr[0]) == 0)
                                //{
                            //}
                            if (selectcm.Count <= 0)
                            {
                                RGL2 cd = new RGL2()
                                {
                                    LOT_ = dr[0].ToString(),
                                    STYLE_ = dr[1].ToString(),
                                    ART = dr[2].ToString(),
                                    COLOR = dr[3].ToString(),
                                    COLORName = dr[4].ToString(),
                                    shangyi_kuzi = dr[5].ToString(),
                                    C34R = dr[6].ToString(),
                                    C36R = dr[7].ToString(),
                                    C38R= dr[8].ToString(),
                                    C40R = dr[9].ToString(),
                                    C42R = dr[10].ToString(),
                                    C44R = dr[11].ToString(),
                                    C46R = dr[12].ToString(),
                                    C48R = dr[13].ToString(),
                                    C50R = dr[14].ToString(),
                                    C52R = dr[15].ToString(),
                                    C54R = dr[16].ToString(),
                                    C56R = dr[17].ToString(),
                                    C58R= dr[18].ToString(),
                                    C60R = dr[19].ToString(),
                                    C62R = dr[20].ToString(),
                                    C36L = dr[21].ToString(),
                                    C38L = dr[22].ToString(),
                                    C40L = dr[23].ToString(),
                                    C42L = dr[24].ToString(),
                                    C44L = dr[25].ToString(),
                                    C46L = dr[26].ToString(),
                                    C48L = dr[27].ToString(),
                                    C50L = dr[28].ToString(),
                                    C52L = dr[29].ToString(),
                                    C54L= dr[30].ToString(),
                                    C56L = dr[31].ToString(),
                                    C58L = dr[32].ToString(),
                                    C60L = dr[33].ToString(),
                                    C62L = dr[34].ToString(),
                                    C34S = dr[35].ToString(),
                                    C36S = dr[36].ToString(),
                                    C38S = dr[37].ToString(),
                                    C40S = dr[38].ToString(),
                                    C42S = dr[39].ToString(),
                                    C44S = dr[40].ToString(),
                                    C46S = dr[41].ToString(),
                                    Sub_Total = dr[42].ToString()
                                };
                                npe.RGL2.Add(cd);
                            }
                            else
                            {
                                string lot = dr[0].ToString();
                                var select = from cd in npe.RGL2
                                             where cd.LOT_.Equals(lot)
                                             select cd;
                                var target = select.FirstOrDefault<RGL2>();
                                    target.LOT_ = dr[0].ToString();
                                    target.STYLE_ = dr[1].ToString();
                                    target.ART = dr[2].ToString();
                                    target.COLOR = dr[3].ToString();
                                    target.COLORName = dr[4].ToString();
                                    target.shangyi_kuzi = dr[5].ToString();
                                    target.C34R = dr[6].ToString();
                                    target.C36R = dr[7].ToString();
                                    target.C38R= dr[8].ToString();
                                    target.C40R = dr[9].ToString();
                                    target.C42R = dr[10].ToString();
                                    target.C44R = dr[11].ToString();
                                    target.C46R = dr[12].ToString();
                                    target.C48R = dr[13].ToString();
                                    target.C50R = dr[14].ToString();
                                    target.C52R = dr[15].ToString();
                                    target.C54R = dr[16].ToString();
                                    target.C56R = dr[17].ToString();
                                    target.C58R= dr[18].ToString();
                                    target.C60R = dr[19].ToString();
                                    target.C62R = dr[20].ToString();
                                    target.C36L = dr[21].ToString();
                                    target.C38L = dr[22].ToString();
                                    target.C40L = dr[23].ToString();
                                    target.C42L = dr[24].ToString();
                                    target.C44L = dr[25].ToString();
                                    target.C46L = dr[26].ToString();
                                    target.C48L = dr[27].ToString();
                                    target.C50L = dr[28].ToString();
                                    target.C52L = dr[29].ToString();
                                    target.C54L= dr[30].ToString();
                                    target.C56L = dr[31].ToString();
                                    target.C58L = dr[32].ToString();
                                    target.C60L = dr[33].ToString();
                                    target.C62L = dr[34].ToString();
                                    target.C34S = dr[35].ToString();
                                    target.C36S = dr[36].ToString();
                                    target.C38S = dr[37].ToString();
                                    target.C40S = dr[38].ToString();
                                    target.C42S = dr[39].ToString();
                                    target.C44S = dr[40].ToString();
                                    target.C46S = dr[41].ToString();
                                    target.Sub_Total = dr[42].ToString();
                            }
                        }
                        npe.SaveChanges();
                    }
                }
                public void InsertChima3(DataTable dt, string biaogeName)
                {
                    using (nemanpingEntities3 npe = new nemanpingEntities3())
                    {
                        List<clsBuiness.SLIM> cmdp = SelectChiMaDapei3();
                        foreach (DataRow dr in dt.Rows)
                        {
                            List<clsBuiness.SLIM> selectcm = new List<SLIM>();
                            //if (dr[0].ToString() != "")
                            //{
                            selectcm = cmdp.FindAll(f => f.LOT_ == dr[0].ToString());
                            //if (dr[0] is DBNull || Convert.ToInt32(dr[0]) == 0)
                            //{
                            //}
                            if (selectcm.Count <= 0)
                            {
                                SLIM cd = new SLIM()
                                {
                                    LOT_ = dr[0].ToString(),
                                    STYLE_ = dr[1].ToString(),
                                    ART = dr[2].ToString(),
                                    COLOR = dr[3].ToString(),
                                    COLOR_ = dr[4].ToString(),
                                    shangyi_kuzi = dr[5].ToString(),
                                    C34R = dr[6].ToString(),
                                    C36R = dr[7].ToString(),
                                    C38R = dr[8].ToString(),
                                    C40R = dr[9].ToString(),
                                    C42R = dr[10].ToString(),
                                    C44R = dr[11].ToString(),
                                    C46R = dr[12].ToString(),
                                    C48R = dr[13].ToString(),
                                    C36L = dr[14].ToString(),
                                    C38L = dr[15].ToString(),
                                    C40L = dr[16].ToString(),
                                    C42L = dr[17].ToString(),
                                    C44L = dr[18].ToString(),
                                    C46L = dr[19].ToString(),
                                    C48L = dr[20].ToString(),
                                    C34S = dr[21].ToString(),
                                    C36S = dr[22].ToString(),
                                    C38S = dr[23].ToString(),
                                    C40S = dr[24].ToString(),
                                    C42S = dr[25].ToString(),
                                    C44S = dr[26].ToString(),
                                    C46S = dr[27].ToString(),
                                    Sub_Total = dr[28].ToString()
                                };
                                npe.SLIM.Add(cd);
                            }
                            else
                            {
                                string lot = dr[0].ToString();
                                var select = from cd in npe.SLIM
                                             where cd.LOT_.Equals(lot)
                                             select cd;
                                var target = select.FirstOrDefault<SLIM>();
                                target.LOT_ = dr[0].ToString();
                                target.STYLE_ = dr[1].ToString();
                                target.ART = dr[2].ToString();
                                target.COLOR = dr[3].ToString();
                                target.COLOR_ = dr[4].ToString();
                                target.shangyi_kuzi = dr[5].ToString();
                                target.C34R = dr[6].ToString();
                                target.C36R = dr[7].ToString();
                                target.C38R = dr[8].ToString();
                                target.C40R = dr[9].ToString();
                                target.C42R = dr[10].ToString();
                                target.C44R = dr[11].ToString();
                                target.C46R = dr[12].ToString();
                                target.C48R = dr[13].ToString();
                                target.C36L = dr[14].ToString();
                                target.C38L = dr[15].ToString();
                                target.C40L = dr[16].ToString();
                                target.C42L = dr[17].ToString();
                                target.C44L = dr[18].ToString();
                                target.C46L = dr[19].ToString();
                                target.C48L = dr[20].ToString();
                                target.C34S = dr[21].ToString();
                                target.C36S = dr[22].ToString();
                                target.C38S = dr[23].ToString();
                                target.C40S = dr[24].ToString();
                                target.C42S = dr[25].ToString();
                                target.C44S = dr[26].ToString();
                                target.C46S = dr[27].ToString();
                                target.Sub_Total = dr[28].ToString();
                            }
                        }
                        npe.SaveChanges();
                    }
                }
                public void InsertChima4(DataTable dt, string biaogeName)
                {
                    using (nemanpingEntities3 npe = new nemanpingEntities3())
                    {
                        List<clsBuiness.RGLJ> cmdp = SelectChiMaDapei4();
                        foreach (DataRow dr in dt.Rows)
                        {
                            List<clsBuiness.RGLJ> selectcm = new List<RGLJ>();
                            //if (dr[0].ToString() != "")
                            //{
                            selectcm = cmdp.FindAll(f => f.LOT_ == dr[0].ToString());
                            //if (dr[0] is DBNull || Convert.ToInt32(dr[0]) == 0)
                            //{
                            //}
                            if (selectcm.Count <= 0)
                            {
                                RGLJ cd = new RGLJ()
                                {
                                    LOT_ = dr[0].ToString(),
                                    STYLE_ = dr[1].ToString(),
                                    ART = dr[2].ToString(),
                                    COLOR = dr[3].ToString(),
                                    COLOR_ = dr[4].ToString(),
                                    shangyi = dr[5].ToString(),
                                    C34R = dr[6].ToString(),
                                    C36R = dr[7].ToString(),
                                    C38R = dr[8].ToString(),
                                    C40R = dr[9].ToString(),
                                    C42R = dr[10].ToString(),
                                    C44R = dr[11].ToString(),
                                    C46R = dr[12].ToString(),
                                    C48R = dr[13].ToString(),
                                    C50R = dr[14].ToString(),
                                    C52R = dr[15].ToString(),
                                    C54R = dr[16].ToString(),
                                    C56R = dr[17].ToString(),
                                    C58R = dr[18].ToString(),
                                    C60R = dr[19].ToString(),
                                    C62R = dr[20].ToString(),
                                    C36L = dr[21].ToString(),
                                    C38L = dr[22].ToString(),
                                    C40L = dr[23].ToString(),
                                    C42L = dr[24].ToString(),
                                    C44L = dr[25].ToString(),
                                    C46L = dr[26].ToString(),
                                    C48L = dr[27].ToString(),
                                    C50L = dr[28].ToString(),
                                    C52L = dr[29].ToString(),
                                    C54L = dr[30].ToString(),
                                    C56L = dr[31].ToString(),
                                    C58L = dr[32].ToString(),
                                    C60L = dr[33].ToString(),
                                    C62L = dr[34].ToString(),
                                    C34S = dr[35].ToString(),
                                    C36S = dr[36].ToString(),
                                    C38S = dr[37].ToString(),
                                    C40S = dr[38].ToString(),
                                    C42S = dr[39].ToString(),
                                    C44S = dr[40].ToString(),
                                    C46S = dr[41].ToString(),
                                    Sub_Total = dr[42].ToString()
                                };
                                npe.RGLJ.Add(cd);
                            }
                            else
                            {
                                string lot = dr[0].ToString();
                                var select = from cd in npe.RGLJ
                                             where cd.LOT_.Equals(lot)
                                             select cd;
                                var target = select.FirstOrDefault<RGLJ>();
                                target.LOT_ = dr[0].ToString();
                                target.STYLE_ = dr[1].ToString();
                                target.ART = dr[2].ToString();
                                target.COLOR = dr[3].ToString();
                                target.COLOR_ = dr[4].ToString();
                                target.shangyi = dr[5].ToString();
                                target.C34R = dr[6].ToString();
                                target.C36R = dr[7].ToString();
                                target.C38R = dr[8].ToString();
                                target.C40R = dr[9].ToString();
                                target.C42R = dr[10].ToString();
                                target.C44R = dr[11].ToString();
                                target.C46R = dr[12].ToString();
                                target.C48R = dr[13].ToString();
                                target.C50R = dr[14].ToString();
                                target.C52R = dr[15].ToString();
                                target.C54R = dr[16].ToString();
                                target.C56R = dr[17].ToString();
                                target.C58R = dr[18].ToString();
                                target.C60R = dr[19].ToString();
                                target.C62R = dr[20].ToString();
                                target.C36L = dr[21].ToString();
                                target.C38L = dr[22].ToString();
                                target.C40L = dr[23].ToString();
                                target.C42L = dr[24].ToString();
                                target.C44L = dr[25].ToString();
                                target.C46L = dr[26].ToString();
                                target.C48L = dr[27].ToString();
                                target.C50L = dr[28].ToString();
                                target.C52L = dr[29].ToString();
                                target.C54L = dr[30].ToString();
                                target.C56L = dr[31].ToString();
                                target.C58L = dr[32].ToString();
                                target.C60L = dr[33].ToString();
                                target.C62L = dr[34].ToString();
                                target.C34S = dr[35].ToString();
                                target.C36S = dr[36].ToString();
                                target.C38S = dr[37].ToString();
                                target.C40S = dr[38].ToString();
                                target.C42S = dr[39].ToString();
                                target.C44S = dr[40].ToString();
                                target.C46S = dr[41].ToString();
                                target.Sub_Total = dr[42].ToString();
                            }
                        }
                        npe.SaveChanges();
                    }
                }
                public void InsertChima5(DataTable dt, string biaogeName)
                {
                    using (nemanpingEntities3 npe = new nemanpingEntities3())
                    {
                        List<clsBuiness.D_PANT> cmdp = SelectChiMaDapei5();
                        foreach (DataRow dr in dt.Rows)
                        {
                            List<clsBuiness.D_PANT> selectcm = new List<D_PANT>();
                            //if (dr[0].ToString() != "")
                            //{
                            selectcm = cmdp.FindAll(f => f.LOT_ == dr[0].ToString());
                            //if (dr[0] is DBNull || Convert.ToInt32(dr[0]) == 0)
                            //{
                            //}
                            if (selectcm.Count <= 0)
                            {
                                if (dr[0] != null && !dr[0].ToString().Equals(string.Empty))
                                {
                                    D_PANT cd = new D_PANT()
                                    {
                                        LOT_ = dr[0].ToString(),
                                        STYLE_ = dr[1].ToString(),
                                        ART = dr[2].ToString(),
                                        COLOR = dr[3].ToString(),
                                        COLORName = dr[4].ToString(),
                                        yaowei = dr[5].ToString(),
                                        C30W_R_30L = dr[6].ToString(),
                                        C30W_L_32L = dr[7].ToString(),
                                        C32W_R_30L = dr[8].ToString(),
                                        C32W_L_32L = dr[9].ToString(),
                                        C34W_S_38L = dr[10].ToString(),
                                        C34W_S_39L = dr[11].ToString(),
                                        C34W_R_30L = dr[12].ToString(),
                                        C34W_L_32L = dr[13].ToString(),
                                        C34W_L_34L = dr[14].ToString(),
                                        C36W_S_28L = dr[15].ToString(),
                                        C36W_S_29L = dr[16].ToString(),
                                        C36W_R_30L = dr[17].ToString(),
                                        C36W_R_31L = dr[18].ToString(),
                                        C38W_S_28L = dr[19].ToString(),
                                        C38W_R_30L = dr[20].ToString(),
                                        C38W_R_31L = dr[21].ToString(),
                                        C38W_L_32L = dr[22].ToString(),
                                        C38W_L_34L = dr[23].ToString(),
                                        C40W_S_28L = dr[24].ToString(),
                                        C40W_S_29L = dr[25].ToString(),
                                        C40W_R_30L = dr[26].ToString(),
                                        C40W_R_31L = dr[27].ToString(),
                                        C40W_L_32L = dr[28].ToString(),
                                        C40W_L_34L = dr[29].ToString(),
                                        C42W_R_30L = dr[30].ToString(),
                                        C42W_L_32L = dr[31].ToString(),
                                        C42W_L_34L = dr[32].ToString(),
                                        C44W_R_30L = dr[33].ToString(),
                                        C44W_L_32L = dr[34].ToString(),
                                        C44W_L_34L = dr[35].ToString(),
                                        C46W_R_30L = dr[36].ToString(),
                                        C46W_L_32L = dr[37].ToString(),
                                        C48W_R_30L = dr[38].ToString(),
                                        C48W_L_32L = dr[39].ToString(),
                                        C50W_L_32L = dr[40].ToString(),
                                        Sub_Total = dr[41].ToString(),
                                        //Sub_Total = dr[43].ToString()
                                    };
                                    npe.D_PANT.Add(cd);
                                }
                            }
                            else
                            {
                                string lot = dr[0].ToString();
                                var select = from cd in npe.D_PANT
                                             where cd.LOT_.Equals(lot)
                                             select cd;
                                var target = select.FirstOrDefault<D_PANT>();
                                target.LOT_ = dr[0].ToString();
                                target.STYLE_ = dr[1].ToString();
                                target.ART = dr[2].ToString();
                                target.COLOR = dr[3].ToString();
                                target.COLORName = dr[4].ToString();
                                target.yaowei = dr[5].ToString();
                                target.C30W_R_30L = dr[6].ToString();
                                target.C30W_L_32L = dr[7].ToString();
                                target.C32W_R_30L = dr[8].ToString();
                                target.C32W_L_32L = dr[9].ToString();
                                target.C34W_S_38L = dr[10].ToString();
                                target.C34W_S_39L = dr[11].ToString();
                                target.C34W_R_30L = dr[12].ToString();
                                target.C34W_L_32L = dr[13].ToString();
                                target.C34W_L_34L = dr[14].ToString();
                                target.C36W_S_28L = dr[15].ToString();
                                target.C36W_S_29L = dr[16].ToString();
                                target.C36W_R_30L = dr[17].ToString();
                                target.C36W_R_31L = dr[18].ToString();
                                target.C38W_S_28L = dr[19].ToString();
                                target.C38W_R_30L = dr[20].ToString();
                                target.C38W_R_31L = dr[21].ToString();
                                target.C38W_L_32L = dr[22].ToString();
                                target.C38W_L_34L = dr[23].ToString();
                                target.C40W_S_28L = dr[24].ToString();
                                target.C40W_S_29L = dr[25].ToString();
                                target.C40W_R_30L = dr[26].ToString();
                                target.C40W_R_31L = dr[27].ToString();
                                target.C40W_L_32L = dr[28].ToString();
                                target.C40W_L_34L = dr[29].ToString();
                                target.C42W_R_30L = dr[30].ToString();
                                target.C42W_L_32L = dr[31].ToString();
                                target.C42W_L_34L = dr[32].ToString();
                                target.C44W_R_30L = dr[33].ToString();
                                target.C44W_L_32L = dr[34].ToString();
                                target.C44W_L_34L = dr[35].ToString();
                                target.C46W_R_30L = dr[36].ToString();
                                target.C46W_L_32L = dr[37].ToString();
                                target.C48W_R_30L = dr[38].ToString();
                                target.C48W_L_32L = dr[39].ToString();
                                target.C50W_L_32L = dr[40].ToString();
                                target.Sub_Total = dr[41].ToString();  
                            }
                        }
                        npe.SaveChanges();
                    }
                }
                public void InsertChima6(DataTable dt, string biaogeName)
                {
                    using (nemanpingEntities3 npe = new nemanpingEntities3())
                    {
                        List<clsBuiness.C_PANT> cmdp = SelectChiMaDapei6();
                        foreach (DataRow dr in dt.Rows)
                        {
                            List<clsBuiness.C_PANT> selectcm = new List<C_PANT>();
                            //if (dr[0].ToString() != "")
                            //{
                            selectcm = cmdp.FindAll(f => f.LOT_ == dr[0].ToString());
                            //if (dr[0] is DBNull || Convert.ToInt32(dr[0]) == 0)
                            //{
                            //}
                            if (selectcm.Count <= 0)
                            {
                                if (dr[0] != null && !dr[0].ToString().Equals(string.Empty))
                                {
                                    C_PANT cd = new C_PANT()
                                    {
                                        LOT_ = dr[0].ToString(),
                                        STYLE_ = dr[1].ToString(),
                                        ART = dr[2].ToString(),
                                        COLOR = dr[3].ToString(),
                                        COLORName = dr[4].ToString(),
                                        yaowei = dr[5].ToString(),
                                        C30W_29L = dr[6].ToString(),
                                        C30W_30L = dr[7].ToString(),
                                        C30W_32L = dr[8].ToString(),
                                        C31W_30L = dr[9].ToString(),
                                        C31W_32L = dr[10].ToString(),
                                        C32W_28L = dr[11].ToString(),
                                        C32W_30L = dr[12].ToString(),
                                        C32W_32L = dr[13].ToString(),
                                        C33W_29L = dr[14].ToString(),
                                        C33W_30L = dr[15].ToString(),
                                        C33W_32L = dr[16].ToString(),
                                        C33W_34L = dr[17].ToString(),
                                        C34W_29L = dr[18].ToString(),
                                        C34W_30L = dr[19].ToString(),
                                        C34W_31L = dr[20].ToString(),
                                        C34W_32L = dr[21].ToString(),
                                        C34W_34L = dr[22].ToString(),
                                        C36W_29L = dr[23].ToString(),
                                        C36W_30L = dr[24].ToString(),
                                        C36W_32L = dr[25].ToString(),
                                        C36W_34L = dr[26].ToString(),
                                        C38W_29L = dr[27].ToString(),
                                        C38W_30L = dr[28].ToString(),
                                        C38W_32L = dr[29].ToString(),
                                        C38W_34L = dr[30].ToString(),
                                        C40W_28L = dr[31].ToString(),
                                        C40W_30L = dr[32].ToString(),
                                        C40W_32L = dr[33].ToString(),
                                        C40W_34L = dr[34].ToString(),
                                        C42W_30L = dr[35].ToString(),
                                        C42W_32L = dr[36].ToString(),
                                        C42W_34L = dr[37].ToString(),
                                        C44W_29L = dr[38].ToString(),
                                        C44W_30L = dr[39].ToString(),
                                        C44W_32L = dr[40].ToString(),
                                        Sub_Total = dr[41].ToString()
                                    };
                                    npe.C_PANT.Add(cd);
                                }
                            }
                            else
                            {
                                string lot = dr[0].ToString();
                                var select = from cd in npe.C_PANT
                                             where cd.LOT_.Equals(lot)
                                             select cd;
                                var target = select.FirstOrDefault<C_PANT>();
                                target.LOT_ = dr[0].ToString();
                                target.STYLE_ = dr[1].ToString();
                                target.ART = dr[2].ToString();
                                target.COLOR = dr[3].ToString();
                                target.COLORName = dr[4].ToString();
                                target.yaowei = dr[5].ToString();
                                target.C30W_29L = dr[6].ToString();
                                target.C30W_30L = dr[7].ToString();
                                target.C30W_32L = dr[8].ToString();
                                target.C31W_30L = dr[9].ToString();
                                target.C31W_32L = dr[10].ToString();
                                target.C32W_28L = dr[11].ToString();
                                target.C32W_30L = dr[12].ToString();
                                target.C32W_32L = dr[13].ToString();
                                target.C33W_29L = dr[14].ToString();
                                target.C33W_30L = dr[15].ToString();
                                target.C33W_32L = dr[16].ToString();
                                target.C33W_34L = dr[17].ToString();
                                target.C34W_29L = dr[18].ToString();
                                target.C34W_30L = dr[19].ToString();
                                target.C34W_31L = dr[20].ToString();
                                target.C34W_32L = dr[21].ToString();
                                target.C34W_34L = dr[22].ToString();
                                target.C36W_29L = dr[23].ToString();
                                target.C36W_30L = dr[24].ToString();
                                target.C36W_32L = dr[25].ToString();
                                target.C36W_34L = dr[26].ToString();
                                target.C38W_29L = dr[27].ToString();
                                target.C38W_30L = dr[28].ToString();
                                target.C38W_32L = dr[29].ToString();
                                target.C38W_34L = dr[30].ToString();
                                target.C40W_28L = dr[31].ToString();
                                target.C40W_30L = dr[32].ToString();
                                target.C40W_32L = dr[33].ToString();
                                target.C40W_34L = dr[34].ToString();
                                target.C42W_30L = dr[35].ToString();
                                target.C42W_32L = dr[36].ToString();
                                target.C42W_34L = dr[37].ToString();
                                target.C44W_29L = dr[38].ToString();
                                target.C44W_30L = dr[39].ToString();
                                target.C44W_32L = dr[40].ToString();
                                target.Sub_Total = dr[41].ToString();
                            }
                        }
                        npe.SaveChanges();
                    }
                }


            #endregion

            #region 读取 EXCEL尺码搭配表
                public List<ChiMa_Dapeibiao> ReaderChiMaDapei(string fileName) 
                {
                    List<ChiMa_Dapeibiao> list= new List<ChiMa_Dapeibiao>();
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
                    {
                        WorkbookPart wbPart = document.WorkbookPart;
                        List<Sheet> sheets = wbPart.Workbook.Descendants<Sheet>().ToList();
                        //var versionSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                        var prosheet = sheets.Find(f => f.Name.Value.Equals("RGL1"));
                        WorksheetPart worksheetPart = (WorksheetPart)wbPart.GetPartById(prosheet.Id);
                        int rowindex = 0;
                        foreach (Row row in worksheetPart.Worksheet.Descendants<Row>())
                        {
                            if (rowindex < 2)
                            {
                                rowindex++;
                                continue;
                            }
                            ChiMa_Dapeibiao target = new ChiMa_Dapeibiao();
                            foreach (Cell cell in row)
                            {
                                string rev = cell.CellReference.Value;
                                if (rev.StartsWith("A")) 
                                {
                                    target.LOT__面料 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("B"))
                                {
                                    target.STYLE_款式 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("C"))
                                {
                                    target.ART_货号 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("D"))
                                {
                                    target.COLOR_颜色 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("E"))
                                {
                                    target.COLOR__颜色编号 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("F"))
                                {
                                    target.JACKET_上衣_PANT_裤子 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("G"))
                                {
                                    target.C34R_28 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("H"))
                                {
                                    target.C36R_30 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("I"))
                                {
                                    target.C38R_32 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("J"))
                                {
                                    target.C40R___34 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("K"))
                                {
                                    target.C42R_36 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("L"))
                                {
                                    target.C44R_38 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("M"))
                                {
                                    target.C46R_40 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("N"))
                                {
                                    target.C48R_42 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("O"))
                                {
                                    target.C50R_44 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("P"))
                                {
                                    target.C52R_46 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("Q"))
                                {
                                    target.C54R_48 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("R"))
                                {
                                    target.C56R_50 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("S"))
                                {
                                    target.C58R_52 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("T"))
                                {
                                    target.C60R_54 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("U"))
                                {
                                    target.C62R_56 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("V"))
                                {
                                    target.C36L_30 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("W"))
                                {
                                    target.C38L_32 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("X"))
                                {
                                    target.C40L_34 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("Y"))
                                {
                                    target.C42L_36 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("Z"))
                                {
                                    target.C44L_38 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AA"))
                                {
                                    target.C46L_40 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AB"))
                                {
                                    target.C48L_42 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AC"))
                                {
                                    target.C50L_44 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AD"))
                                {
                                    target.C52L_46 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AE"))
                                {
                                    target.C54L_48 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AF"))
                                {
                                    target.C56L_50 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AG"))
                                {
                                    target.C58L_52 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AH"))
                                {
                                    target.C60L_54 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AI"))
                                {
                                    target.C62L_56 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AJ"))
                                {
                                    target.C34S_28 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AK"))
                                {
                                    target.C36S_30 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AL"))
                                {
                                    target.C38S_32 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AM"))
                                {
                                    target.C40S_34 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AN"))
                                {
                                    target.C42S_36 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AO"))
                                {
                                    target.C44S_38 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AP"))
                                {
                                    target.C46S_40 = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AQ"))
                                {
                                    target.DingdanHeji = GetCellValue(wbPart, cell);
                                }
                            }
                            list.Add(target);
                        }
                    }
                    return list;
                }
                public List<RGL2> ReaderChiMaDapei2(string fileName)
                {
                    List<RGL2> list = new List<RGL2>();
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
                    {
                        WorkbookPart wbPart = document.WorkbookPart;
                        List<Sheet> sheets = wbPart.Workbook.Descendants<Sheet>().ToList();
                        var prosheet = sheets.Find(f => f.Name.Value.Equals("RGL2"));
                        WorksheetPart worksheetPart = (WorksheetPart)wbPart.GetPartById(prosheet.Id);

                        int rowindex = 0;
                        foreach (Row row in worksheetPart.Worksheet.Descendants<Row>())
                        {
                            if (rowindex < 2)
                            {
                                rowindex++;
                                continue;
                            }
                            RGL2 target = new RGL2();
                            foreach (Cell cell in row)
                            {
                                string rev = cell.CellReference.Value;
                                if (rev.StartsWith("A"))
                                {
                                    target.LOT_ = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("B"))
                                {
                                    target.STYLE_ = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("C"))
                                {
                                    target.ART = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("D"))
                                {
                                    target.COLOR = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("E"))
                                {
                                    target.COLORName = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("F"))
                                {
                                    target.shangyi_kuzi = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("G"))
                                {
                                    target.C34R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("H"))
                                {
                                    target.C36R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("I"))
                                {
                                    target.C38R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("J"))
                                {
                                    target.C40R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("K"))
                                {
                                    target.C42R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("L"))
                                {
                                    target.C44R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("M"))
                                {
                                    target.C46R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("N"))
                                {
                                    target.C48R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("O"))
                                {
                                    target.C50R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("P"))
                                {
                                    target.C52R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("Q"))
                                {
                                    target.C54R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("R"))
                                {
                                    target.C56R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("S"))
                                {
                                    target.C58R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("T"))
                                {
                                    target.C60R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("U"))
                                {
                                    target.C62R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("V"))
                                {
                                    target.C36L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("W"))
                                {
                                    target.C38L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("X"))
                                {
                                    target.C40L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("Y"))
                                {
                                    target.C42L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("Z"))
                                {
                                    target.C44L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AA"))
                                {
                                    target.C46L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AB"))
                                {
                                    target.C48L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AC"))
                                {
                                    target.C50L= GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AD"))
                                {
                                    target.C52L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AE"))
                                {
                                    target.C54L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AF"))
                                {
                                    target.C56L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AG"))
                                {
                                    target.C58L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AH"))
                                {
                                    target.C60L= GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AI"))
                                {
                                    target.C62L= GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AJ"))
                                {
                                    target.C34S = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AK"))
                                {
                                    target.C36S = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AL"))
                                {
                                    target.C38S = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AM"))
                                {
                                    target.C40S = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AN"))
                                {
                                    target.C42S = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AO"))
                                {
                                    target.C44S = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AP"))
                                {
                                    target.C46S= GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AQ"))
                                {
                                    target.Sub_Total = GetCellValue(wbPart, cell);
                                }
                            }
                            list.Add(target);
                        }
                    }
                    return list;
                }
                public List<SLIM> ReaderChiMaDapei3(string fileName)
                {
                    List<SLIM> list = new List<SLIM>();
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
                    {
                        WorkbookPart wbPart = document.WorkbookPart;
                        List<Sheet> sheets = wbPart.Workbook.Descendants<Sheet>().ToList();
                        //var prosheet = sheets.FindAll(f => f.Name.Equals("SLIM"));
                        //WorksheetPart worksheetPart = (WorksheetPart)wbPart.GetPartById(prosheet[0].Id);
                        var prosheet = sheets.Find(f => f.Name.Value.Equals("SLIM"));
                        WorksheetPart worksheetPart = (WorksheetPart)wbPart.GetPartById(prosheet.Id);

                        int rowindex = 0;
                        foreach (Row row in worksheetPart.Worksheet.Descendants<Row>())
                        {
                            if (rowindex < 2)
                            {
                                rowindex++;
                                continue;
                            }
                            SLIM target = new SLIM();
                            foreach (Cell cell in row)
                            {
                                string rev = cell.CellReference.Value;
                                if (rev.StartsWith("A"))
                                {
                                    target.LOT_ = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("B"))
                                {
                                    target.STYLE_= GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("C"))
                                {
                                    target.ART = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("D"))
                                {
                                    target.COLOR = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("E"))
                                {
                                    target.COLOR_ = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("F"))
                                {
                                    target.shangyi_kuzi = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("G"))
                                {
                                    target.C34R= GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("H"))
                                {
                                    target.C36R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("I"))
                                {
                                    target.C38R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("J"))
                                {
                                    target.C40R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("K"))
                                {
                                    target.C42R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("L"))
                                {
                                    target.C44R= GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("M"))
                                {
                                    target.C46R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("N"))
                                {
                                    target.C48R = GetCellValue(wbPart, cell);
                                }
                               
                                if (rev.StartsWith("O"))
                                {
                                    target.C36L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("P"))
                                {
                                    target.C38L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("Q"))
                                {
                                    target.C40L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("R"))
                                {
                                    target.C42L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("S"))
                                {
                                    target.C44L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("T"))
                                {
                                    target.C46L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("U"))
                                {
                                    target.C48L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("V"))
                                {
                                    target.C34S = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("W"))
                                {
                                    target.C36S = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("X"))
                                {
                                    target.C38S= GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("Y"))
                                {
                                    target.C40S = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("Z"))
                                {
                                    target.C42S= GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AA"))
                                {
                                    target.C44S = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AB"))
                                {
                                    target.C46S = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AC"))
                                {
                                    target.Sub_Total = GetCellValue(wbPart, cell);
                                }
                            }
                            list.Add(target);
                        }
                    }
                    return list;
                }
                public List<RGLJ> ReaderChiMaDapei4(string fileName)
                {
                    List<RGLJ> list = new List<RGLJ>();
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
                    {
                        WorkbookPart wbPart = document.WorkbookPart;
                        List<Sheet> sheets = wbPart.Workbook.Descendants<Sheet>().ToList();
                        //var prosheet = sheets.FindAll(f => f.Name.Equals("RGLJ"));
                        //WorksheetPart worksheetPart = (WorksheetPart)wbPart.GetPartById(prosheet[0].Id);
                        var prosheet = sheets.Find(f => f.Name.Value.Equals("RGLJ"));
                        WorksheetPart worksheetPart = (WorksheetPart)wbPart.GetPartById(prosheet.Id);

                        int rowindex = 0;
                        foreach (Row row in worksheetPart.Worksheet.Descendants<Row>())
                        {
                            if (rowindex < 2)
                            {
                                rowindex++;
                                continue;
                            }
                            RGLJ target = new RGLJ();
                            foreach (Cell cell in row)
                            {
                                string rev = cell.CellReference.Value;
                                if (rev.StartsWith("A"))
                                {
                                    target.LOT_ = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("B"))
                                {
                                    target.STYLE_ = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("C"))
                                {
                                    target.ART = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("D"))
                                {
                                    target.COLOR = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("E"))
                                {
                                    target.COLOR_ = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("F"))
                                {
                                    target.shangyi = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("G"))
                                {
                                    target.C34R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("H"))
                                {
                                    target.C36R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("I"))
                                {
                                    target.C38R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("J"))
                                {
                                    target.C40R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("K"))
                                {
                                    target.C42R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("L"))
                                {
                                    target.C44R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("M"))
                                {
                                    target.C46R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("N"))
                                {
                                    target.C48R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("O"))
                                {
                                    target.C50R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("P"))
                                {
                                    target.C52R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("Q"))
                                {
                                    target.C54R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("R"))
                                {
                                    target.C56R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("S"))
                                {
                                    target.C58R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("T"))
                                {
                                    target.C60R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("U"))
                                {
                                    target.C62R = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("V"))
                                {
                                    target.C36L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("W"))
                                {
                                    target.C38L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("X"))
                                {
                                    target.C40L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("Y"))
                                {
                                    target.C42L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("Z"))
                                {
                                    target.C44L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AA"))
                                {
                                    target.C46L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AB"))
                                {
                                    target.C48L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AC"))
                                {
                                    target.C50L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AD"))
                                {
                                    target.C52L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AE"))
                                {
                                    target.C54L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AF"))
                                {
                                    target.C56L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AG"))
                                {
                                    target.C58L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AH"))
                                {
                                    target.C60L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AI"))
                                {
                                    target.C62L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AJ"))
                                {
                                    target.C34S = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AK"))
                                {
                                    target.C36S = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AL"))
                                {
                                    target.C38S = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AM"))
                                {
                                    target.C40S = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AN"))
                                {
                                    target.C42S = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AO"))
                                {
                                    target.C44S = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AP"))
                                {
                                    target.C46S = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AQ"))
                                {
                                    target.Sub_Total = GetCellValue(wbPart, cell);
                                }
                            }
                            list.Add(target);
                        }
                    }
                    return list;
                }
                public List<D_PANT> ReaderChiMaDapei5(string fileName)
                {
                    List<D_PANT> list = new List<D_PANT>();
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
                    {
                        WorkbookPart wbPart = document.WorkbookPart;
                        List<Sheet> sheets = wbPart.Workbook.Descendants<Sheet>().ToList();
                        //var prosheet = sheets.FindAll(f => f.Name.Equals("D.PANT"));
                        //WorksheetPart worksheetPart = (WorksheetPart)wbPart.GetPartById(prosheet[0].Id);
                        var prosheet = sheets.Find(f => f.Name.Value.Equals("D.PANT"));
                        WorksheetPart worksheetPart = (WorksheetPart)wbPart.GetPartById(prosheet.Id);

                        int rowindex = 0;
                        foreach (Row row in worksheetPart.Worksheet.Descendants<Row>())
                        {
                            if (rowindex < 2)
                            {
                                rowindex++;
                                continue;
                            }
                            D_PANT target = new D_PANT();
                            foreach (Cell cell in row)
                            {
                                string rev = cell.CellReference.Value;
                                if (rev.StartsWith("A"))
                                {
                                    target.LOT_ = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("B"))
                                {
                                    target.STYLE_= GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("C"))
                                {
                                    target.ART = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("D"))
                                {
                                    target.COLOR = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("E"))
                                {
                                    target.COLORName = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("F"))
                                {
                                    target.yaowei = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("G"))
                                {
                                    target.C30W_R_30L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("H"))
                                {
                                    target.C30W_L_32L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("I"))
                                {
                                    target.C32W_R_30L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("J"))
                                {
                                    target.C32W_L_32L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("K"))
                                {
                                    target.C34W_S_38L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("L"))
                                {
                                    target.C34W_S_39L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("M"))
                                {
                                    target.C34W_R_30L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("N"))
                                {
                                    target.C34W_L_32L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("O"))
                                {
                                    target.C34W_L_34L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("P"))
                                {
                                    target.C36W_S_28L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("Q"))
                                {
                                    target.C36W_S_29L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("R"))
                                {
                                    target.C36W_R_30L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("S"))
                                {
                                    target.C36W_R_31L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("T"))
                                {
                                    target.C38W_S_28L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("U"))
                                {
                                    target.C38W_R_30L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("V"))
                                {
                                    target.C38W_R_31L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("W"))
                                {
                                    target.C38W_L_32L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("X"))
                                {
                                    target.C38W_L_34L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("Y"))
                                {
                                    target.C40W_S_28L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("Z"))
                                {
                                    target.C40W_S_29L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AA"))
                                {
                                    target.C40W_R_30L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AB"))
                                {
                                    target.C40W_R_31L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AC"))
                                {
                                    target.C40W_L_32L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AD"))
                                {
                                    target.C40W_L_34L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AE"))
                                {
                                    target.C42W_R_30L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AF"))
                                {
                                    target.C42W_L_32L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AG"))
                                {
                                    target.C42W_L_34L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AH"))
                                {
                                    target.C44W_R_30L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AI"))
                                {
                                    target.C44W_L_32L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AJ"))
                                {
                                    target.C44W_L_34L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AK"))
                                {
                                    target.C46W_R_30L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AL"))
                                {
                                    target.C46W_L_32L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AM"))
                                {
                                    target.C48W_R_30L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AN"))
                                {
                                    target.C48W_L_32L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AO"))
                                {
                                    target.C50W_L_32L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AP"))
                                {
                                    target.Sub_Total = GetCellValue(wbPart, cell);
                                }
                                //if (rev.StartsWith("AQ"))
                                //{
                                //    target.DingdanHeji = GetCellValue(wbPart, cell);
                                //}
                            }
                            list.Add(target);
                        }
                    }
                    return list;
                }
                public List<C_PANT> ReaderChiMaDapei6(string fileName)
                {
                    List<C_PANT> list = new List<C_PANT>();
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
                    {
                        WorkbookPart wbPart = document.WorkbookPart;
                        List<Sheet> sheets = wbPart.Workbook.Descendants<Sheet>().ToList();
                        //var prosheet = sheets.FindAll(f => f.Name.Equals("C.PANT"));
                        //WorksheetPart worksheetPart = (WorksheetPart)wbPart.GetPartById(prosheet[0].Id);
                        var prosheet = sheets.Find(f => f.Name.Value.Equals("C.PANT"));
                        WorksheetPart worksheetPart = (WorksheetPart)wbPart.GetPartById(prosheet.Id);

                        int rowindex = 0;
                        foreach (Row row in worksheetPart.Worksheet.Descendants<Row>())
                        {
                            if (rowindex < 2)
                            {
                                rowindex++;
                                continue;
                            }
                            C_PANT target = new C_PANT();
                            foreach (Cell cell in row)
                            {
                                string rev = cell.CellReference.Value;
                                if (rev.StartsWith("A"))
                                {
                                    target.LOT_ = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("B"))
                                {
                                    target.STYLE_= GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("C"))
                                {
                                    target.ART = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("D"))
                                {
                                    target.COLOR = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("E"))
                                {
                                    target.COLORName = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("F"))
                                {
                                    target.yaowei = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("G"))
                                {
                                    target.C30W_29L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("H"))
                                {
                                    target.C30W_30L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("I"))
                                {
                                    target.C30W_32L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("J"))
                                {
                                    target.C31W_30L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("K"))
                                {
                                    target.C31W_32L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("L"))
                                {
                                    target.C32W_28L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("M"))
                                {
                                    target.C32W_30L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("N"))
                                {
                                    target.C32W_32L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("O"))
                                {
                                    target.C33W_29L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("P"))
                                {
                                    target.C33W_30L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("Q"))
                                {
                                    target.C33W_32L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("R"))
                                {
                                    target.C33W_34L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("S"))
                                {
                                    target.C34W_29L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("T"))
                                {
                                    target.C34W_30L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("U"))
                                {
                                    target.C34W_31L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("V"))
                                {
                                    target.C34W_32L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("W"))
                                {
                                    target.C34W_32L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("X"))
                                {
                                    target.C36W_29L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("Y"))
                                {
                                    target.C36W_30L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("Z"))
                                {
                                    target.C36W_32L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AA"))
                                {
                                    target.C36W_34L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AB"))
                                {
                                    target.C38W_29L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AC"))
                                {
                                    target.C38W_30L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AD"))
                                {
                                    target.C38W_32L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AE"))
                                {
                                    target.C38W_34L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AF"))
                                {
                                    target.C40W_28L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AG"))
                                {
                                    target.C40W_30L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AH"))
                                {
                                    target.C40W_32L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AI"))
                                {
                                    target.C40W_34L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AJ"))
                                {
                                    target.C42W_30L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AK"))
                                {
                                    target.C42W_32L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AL"))
                                {
                                    target.C42W_34L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AM"))
                                {
                                    target.C44W_29L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AN"))
                                {
                                    target.C44W_30L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AO"))
                                {
                                    target.C44W_32L = GetCellValue(wbPart, cell);
                                }
                                if (rev.StartsWith("AP"))
                                {
                                    target.Sub_Total = GetCellValue(wbPart, cell);
                                }
                                //if (rev.StartsWith("AQ"))
                                //{
                                //    target.DingdanHeji = GetCellValue(wbPart, cell);
                                //}
                            }
                            list.Add(target);
                        }
                    }
                    return list;
                }
            #endregion

            #region 删除尺码搭配信息
                public void deleteChiMa(List<int> id) 
                {
                    using (nemanpingEntities3 npe = new nemanpingEntities3())
                    {
                        foreach (int strid in id)
                        {
                            if (strid != null)
                            {
                                if (strid != 0)
                                {
                                    var select = from s in npe.ChiMa_Dapeibiao where s.id == strid select s;
                                    foreach (var item in select)
                                    {
                                        npe.ChiMa_Dapeibiao.Remove(item);
                                    }
                                }
                            }
                        }
                        npe.SaveChanges();
                    }
                
                }
            #endregion

            #region 删除尺码搭配表
                public void deleteChiMaBiao(string id)
                {
                    using (nemanpingEntities3 npe = new nemanpingEntities3())
                    {
                        foreach (int strid in id)
                        {
                            if (strid != null)
                            {
                                if (strid != 0)
                                {
                                    var select = from s in npe.ChiMa_Dapeibiao where s.BiaoName.Equals(id) select s;
                                    foreach (var item in select)
                                    {
                                        npe.ChiMa_Dapeibiao.Remove(item);
                                    }
                                }
                            }
                        }
                        npe.SaveChanges();
                    }

                }

            #endregion
       #endregion

        #region 功能:款式表录入
                #region 查询款式表
                public List<KuanShiBiao> SelectKuanshi()
                {
                    using (nemanpingEntities3 npe = new nemanpingEntities3())
                    {
                        List<KuanShiBiao> list = new List<KuanShiBiao>();

                        var select = from cm in npe.KuanShiBiao
                                     select cm;
                        list = select.ToList<KuanShiBiao>();
                        return list;
                    }

                }
            #endregion

            #region 提交款式表数据
                public void insertKuanShi(DataTable dt)
                {
                    using (nemanpingEntities3 can = new nemanpingEntities3())
                    {
                        foreach (DataRow dr in dt.Rows)
                        {

                            if (dr[0] is DBNull || Convert.ToInt32(dr[0]) == 0)
                            {
                                KuanShiBiao ks = new KuanShiBiao()
                                {
                                    STYLE = dr[1].ToString(),
                                    DESC = dr[2].ToString(),
                                    FABRIC = dr[3].ToString(),
                                    JACKET = dr[4].ToString(),
                                    PANT = dr[5].ToString(),
                                    ShuoMing = dr[7].ToString(),
                                    mark1 = dr[6].ToString(),
                                    mark2 = dr[8].ToString()
                                };
                                can.KuanShiBiao.Add(ks);

                            }
                            else
                            {
                                int id = Convert.ToInt32(dr[0]);
                                var select = from ks in can.KuanShiBiao where ks.Id == id select ks;
                                var target = select.FirstOrDefault<KuanShiBiao>();
                                target.STYLE = dr[1].ToString();
                                target.DESC = dr[2].ToString();
                                target.FABRIC = dr[3].ToString();
                                target.JACKET = dr[4].ToString();
                                target.PANT = dr[5].ToString();
                                target.ShuoMing = dr[7].ToString();
                                target.mark1 = dr[6].ToString();
                                target.mark2 = dr[8].ToString();

                            }
                        }
                        can.SaveChanges();

                    }
                }
            #endregion

            #region 读取 Excel款式表
                public List<KuanShiBiao> readerKuanshi(string fileName)
                {
                    List<KuanShiBiao> list = new List<KuanShiBiao>();
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
                    {
                        WorkbookPart wbPart = document.WorkbookPart;
                        List<Sheet> sheets = wbPart.Workbook.Descendants<Sheet>().ToList();
                        var versionSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                        WorksheetPart worksheetPart = (WorksheetPart)wbPart.GetPartById(versionSheet.Id);
                        int rowindex = 0;
                        int insertpd = 0;
                        int descinsertPd = 0;
                        int mlpd = 0;
                        int jackpd = 0;
                        int pantpd = 0;
                        int smpd = 0;
                        int listinsertpd =0;
                        int vspd = 0;
                        KuanShiBiao ks = new KuanShiBiao();
                        foreach (Row row in worksheetPart.Worksheet.Descendants<Row>())
                        {
                            if (rowindex < 2)
                            {
                                rowindex++;
                                continue;
                            }
                            
                            foreach (Cell cell in row)
                            {
                                
                                string rev = cell.CellReference.Value;
                                
                                if (rev.StartsWith("B"))  
                                {
                                    if (GetCellValue(wbPart, cell).Equals(string.Empty))
                                    {
                                        listinsertpd = 1;
                                        

                                    }
                                    if (insertpd == 1)
                                    {
                                        ks.STYLE = GetCellValue(wbPart, cell);
                                        insertpd = 0;
                                    }
                                }
                                if (rev.StartsWith("C")) 
                                {
                                    if (descinsertPd == 1)
                                    {
                                        ks.DESC = GetCellValue(wbPart, cell);
                                        descinsertPd = 0;
                                    }
                                }
                                
                                if (rev.StartsWith("D")) 
                                {
                                    if (mlpd == 1) 
                                    {
                                        ks.FABRIC = GetCellValue(wbPart, cell);
                                        mlpd = 0;
                                    }
                                    if (jackpd == 1) 
                                    {
                                        ks.JACKET = GetCellValue(wbPart,cell);
                                    }
                                    if (pantpd == 1)
                                    {
                                        ks.PANT = GetCellValue(wbPart, cell);
                                    }
                                    if (vspd == 1)
                                    {
                                        ks.mark1 = ks.mark1 + " " + GetCellValue(wbPart, cell);
                                    }
                                    if (smpd == 1) 
                                    {
                                        ks.ShuoMing = GetCellValue(wbPart, cell);
                                        smpd = 0;
                                    }
                                }
                                if (rev.StartsWith("E")) 
                                {
                                    if (jackpd == 1)
                                    {
                                        ks.JACKET = ks.JACKET+" "+ GetCellValue(wbPart, cell);
                                        jackpd = 0;
                                    }
                                    if (pantpd == 1)
                                    {
                                        ks.PANT =ks.PANT+" "+ GetCellValue(wbPart, cell);
                                        pantpd = 0;
                                    }
                                    if (vspd == 1) 
                                    {
                                        ks.mark1 = ks.mark1 + " " + GetCellValue(wbPart, cell);
                                        pantpd = 0;
                                    }
                                }
                                #region 判断该条数据是否添加
                                if (rev.StartsWith("A"))
                                {
                                    if (!GetCellValue(wbPart, cell).Equals(string.Empty))
                                    {
                                        insertpd = 1;
                                    }
                                }
                                if (GetCellValue(wbPart, cell).Contains("DESC"))
                                {
                                    descinsertPd = 1;
                                }
                                if (GetCellValue(wbPart, cell).Contains("FABRIC 面料成份："))
                                {
                                    mlpd = 1;
                                }
                                if (GetCellValue(wbPart, cell).Contains("JACKET"))
                                {
                                    jackpd = 1;
                                }
                                if (GetCellValue(wbPart, cell).Contains("PANT"))
                                {
                                    pantpd = 1;
                                }
                                if (GetCellValue(wbPart, cell).Contains("说明"))
                                {
                                    smpd = 1;
                                }
                                if (GetCellValue(wbPart, cell).Contains("VEST")) 
                                {
                                    vspd = 1;
                                }

                                #endregion
                            }
                            if (listinsertpd == 1) 
                            {
                                
                                list.Add(ks);
                                ks = new KuanShiBiao();
                                listinsertpd = 0;
                            }
                        }
                        return list;
                    }
                }
            #endregion

            #region 删除款式表
                public void deleteKuanshi(List<int> id)
                {
                    using (nemanpingEntities3 npe = new nemanpingEntities3())
                    {
                        foreach (int strid in id)
                        {
                            if (strid != null)
                            {
                                if (strid != 0)
                                {
                                    var select = from s in npe.KuanShiBiao where s.Id == strid select s;
                                    foreach (var item in select)
                                    {
                                        npe.KuanShiBiao.Remove(item);
                                    }
                                }
                            }
                        }
                        npe.SaveChanges();
                    }

                }
            #endregion
        #endregion

        #region 功能:单耗表录入
            #region 查询单耗表
                public List<DanHao> SelectDanHao(string strwhere) 
                {
                    using (nemanpingEntities3 npe = new nemanpingEntities3())
                    {
                        List<DanHao> list = new List<DanHao>();

                        if (strwhere.Equals(string.Empty))
                        {
                            var select = from cm in npe.DanHao
                                         select cm;
                            list = select.ToList<DanHao>();
                        }
                        else
                        {
                            var select = from cm in npe.DanHao
                                         where cm.CaiDanNo.Equals(strwhere)
                                         select cm;
                            list = select.ToList<DanHao>();
                        }
                        return list;
                    }
                }
            #endregion

            #region 提交单耗表
                public void insertDanhao(DataTable dt) 
                {
                    using(nemanpingEntities3 nep = new nemanpingEntities3())
                    {
                        List<clsBuiness.DanHao> biao = SelectDanHao("");
                        List<clsBuiness.DanHao> dh = new List<DanHao>();
                        foreach (DataRow dr in dt.Rows)
                        {
                            //if(dr[0].ToString() != "" )
                            //{
                                dh = biao.FindAll(f => f.CaiDanNo.Equals(dr[1].ToString()) &&  f.Name.Equals(dr[6].ToString()));
                            //}
                            //if (dr[0] is DBNull || Convert.ToInt32(dr[0]) == 0)
                            //{
                            if(dh.Count <=0){
                                DanHao insets = new DanHao()
                                {
                                    CaiDanNo = dr[1].ToString(),
                                    Style = dr[2].ToString(),
                                    FABRIC_CONTENT = dr[3].ToString(),
                                    DATE = dr[4].ToString(),
                                    JiaGongChang = dr[5].ToString(),
                                    Name = dr[6].ToString(),
                                    HuoHao = dr[7].ToString(),
                                    GuiGe = dr[8].ToString(),
                                    Yanse = dr[9].ToString(),
                                    Danjia = dr[10].ToString(),
                                    DanHao1 = dr[11].ToString(),
                                    Jine = dr[12].ToString(),
                                    BeiZhu = dr[13].ToString(),
                                    ChangShang = dr[14].ToString(),
                                    Heji = dr[15].ToString(),
                                    Type = dr[16].ToString()
                                };

                                nep.DanHao.Add(insets);

                            }
                            else
                            {
                                int id = Convert.ToInt32(dr[0]);
                                string name = dr[6].ToString();
                                string caidan = dr[1].ToString();
                                var select = from sc in nep.DanHao where sc.Name.Equals(name) && sc.CaiDanNo.Equals(caidan) select sc;
                                var target = select.FirstOrDefault<DanHao>();
                                target.Name = dr[6].ToString();
                                target.HuoHao = dr[7].ToString();
                                target.GuiGe = dr[8].ToString();
                                target.Yanse = dr[9].ToString();
                                target.Danjia = dr[10].ToString();
                                target.DanHao1 = dr[11].ToString();
                                target.Jine = dr[12].ToString();
                                target.BeiZhu = dr[13].ToString();
                                target.ChangShang = dr[14].ToString();
                                target.Heji = dr[15].ToString();
                                target.Type = dr[16].ToString();
                            }
                        }
                        nep.SaveChanges();
                    }
                }

            #endregion

            #region 读取EXCEL单耗表
                public List<DanHao> Readerdh(string fileName) 
                {
                    List<DanHao> list = new List<DanHao>();
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
                    {
                        WorkbookPart wbPart = document.WorkbookPart;
                        List<Sheet> sheets = wbPart.Workbook.Descendants<Sheet>().ToList();
                        var versionSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                        WorksheetPart worksheetPart = (WorksheetPart)wbPart.GetPartById(versionSheet.Id);
                        int rowindex = 0;
                        int InsertPd = 0;
                        string jiagongchang="";
                        string caidanhao="";
                        string kuanshi="";
                        string mlcf="";
                        string riqi="";
                        bool addpd = false;
                        bool typepd = false;
                        foreach (Row row in worksheetPart.Worksheet.Descendants<Row>())
                        {
                            if (rowindex < 1)
                            {
                                rowindex++;
                                continue;
                            }
                            DanHao d = new DanHao();
                            foreach (Cell cell in row)
                            {
                                string rev = cell.CellReference.Value;
                                if (InsertPd == 0) 
                                {
                                    if (rev.StartsWith("A") && rev.EndsWith("2"))
                                    {
                                        caidanhao = GetCellValue(wbPart, cell).Split('：')[1];
                                    }

                                    if (rev.StartsWith("A") && rev.EndsWith("3"))
                                    {
                                        kuanshi = GetCellValue(wbPart, cell).Split('：')[1];
                                    }
                                    if (rev.StartsWith("E") && rev.EndsWith("2"))
                                    {
                                        riqi = GetCellValue(wbPart, cell).Split('：')[1];
                                    }

                                    if (rev.StartsWith("E") && rev.EndsWith("3"))
                                    {
                                        jiagongchang = GetCellValue(wbPart, cell).Split('：')[1];
                                    }
                                    if (rev.StartsWith("A") && rev.EndsWith("4"))
                                    {
                                        mlcf = GetCellValue(wbPart, cell).Split('：')[1];
                                        InsertPd++;
                                    }
                                }
                                if (!rev.EndsWith("5")) 
                                {
                                    if (InsertPd >= 1)
                                    {
                                        addpd = true;
                                        d.CaiDanNo = caidanhao;
                                        d.JiaGongChang = jiagongchang;
                                        d.DATE = riqi;
                                        d.Style = kuanshi;
                                        d.FABRIC_CONTENT = mlcf;
                                        if (typepd)
                                        {
                                            if (rev.StartsWith("B"))
                                            {
                                                d.Name = GetCellValue(wbPart, cell);
                                            }
                                            if (rev.StartsWith("C"))
                                            {
                                                d.HuoHao = GetCellValue(wbPart, cell);
                                            }
                                            if (rev.StartsWith("D"))
                                            {
                                                d.GuiGe = GetCellValue(wbPart, cell);
                                            }
                                            if (rev.StartsWith("E"))
                                            {
                                                d.Yanse = GetCellValue(wbPart, cell);
                                            }
                                            if (rev.StartsWith("F"))
                                            {
                                                d.Danjia = GetCellValue(wbPart, cell);
                                            }
                                            if (rev.StartsWith("G"))
                                            {
                                                d.DanHao1 = GetCellValue(wbPart, cell);
                                            }
                                            if (rev.StartsWith("H"))
                                            {
                                                d.Jine = GetCellValue(wbPart, cell);
                                            }
                                            if (rev.StartsWith("I"))
                                            {
                                                d.BeiZhu = GetCellValue(wbPart, cell);
                                            }
                                            if (rev.StartsWith("J"))
                                            {
                                                d.ChangShang = GetCellValue(wbPart, cell);
                                            }
                                            d.Type = "辅料";
                                        }
                                        else
                                        {
                                            if (rev.StartsWith("B"))
                                            {
                                                d.Name = GetCellValue(wbPart, cell);
                                            }
                                            if (rev.StartsWith("C"))
                                            {
                                                d.HuoHao = GetCellValue(wbPart, cell);
                                            }
                                            if (rev.StartsWith("D"))
                                            {
                                                d.GuiGe = GetCellValue(wbPart, cell);
                                            }
                                            if (rev.StartsWith("E"))
                                            {
                                                d.Yanse = GetCellValue(wbPart, cell);
                                            }
                                            if (rev.StartsWith("F"))
                                            {
                                                d.Danjia = GetCellValue(wbPart, cell);
                                            }
                                            if (rev.StartsWith("G"))
                                            {
                                                d.DanHao1 = GetCellValue(wbPart, cell);
                                            }
                                            if (rev.StartsWith("H"))
                                            {
                                                d.Jine = GetCellValue(wbPart, cell);
                                            }
                                            if (rev.StartsWith("I"))
                                            {
                                                d.BeiZhu = GetCellValue(wbPart, cell);
                                            }
                                            if (rev.StartsWith("J"))
                                            {
                                                d.ChangShang = GetCellValue(wbPart, cell);
                                            }
                                            d.Type = "面料";
                                        }
                                        if (GetCellValue(wbPart, cell).Contains("辅料"))
                                        {
                                            typepd = true;
                                        }

                                    }

                                }

                            }
                            if (addpd) 
                            {
                                list.Add(d);
                                addpd = false;
                            }
                            
                        }
                        return list;
                    }
                }
            #endregion

            #region 删除单耗表数据
                public void deleteDanHao(List<int> id)
                {
                    using (nemanpingEntities3 npe = new nemanpingEntities3())
                    {
                        foreach (int strid in id)
                        {
                            if (strid != null)
                            {
                                if (strid != 0)
                                {
                                    var select = from s in npe.DanHao where s.Id == strid select s;
                                    foreach (var item in select)
                                    {
                                        npe.DanHao.Remove(item);
                                    }
                                }
                            }
                        }
                        npe.SaveChanges();
                    }

                }
            #endregion

            #region 删除单耗表
                public void deletDh(string caidanhao) 
                {
                    using (nemanpingEntities3 npe = new nemanpingEntities3())
                    {

                        if (caidanhao != null)
                            {
                                var select = from s in npe.DanHao where s.CaiDanNo.Equals(caidanhao) select s;
                                foreach (var item in select)
                                {
                                    npe.DanHao.Remove(item);
                                }
                            }
                        
                        npe.SaveChanges();
                    }
                }
                #endregion
        #endregion

        #region 功能:配色表录入
                #region 查询配色表
                public List<PeiSe> selectPeise(string strwhere)
                {
                    using (nemanpingEntities3 can = new nemanpingEntities3())
                    {
                        List<PeiSe> sehao = new List<PeiSe>();
                        if (strwhere.Equals(string.Empty))
                        {
                            var select = from s in can.PeiSe select s;
                            sehao = select.ToList<PeiSe>();
                        }
                        else 
                        {
                            var select = from s in can.PeiSe where s.Fabrics == strwhere select s;
                            sehao = select.ToList<PeiSe>();
                        }
                       

                        return sehao;
                    }
                }
            
            #endregion

                #region 提交配色信息
                public void insertPeise(DataTable dt, string Fabrics,string date) 
                {
                    using (nemanpingEntities3 nep = new nemanpingEntities3()) 
                    {
                        List<clsBuiness.PeiSe> peise = new List<PeiSe>();
                        List<clsBuiness.PeiSe> biao = selectPeise("");
                        foreach (DataRow dr in dt.Rows)
                        {
                            if (dr[19].ToString() !="")
                            {
                                peise = biao.FindAll(f => f.Id == Convert.ToInt32(dr[19]) && f.Fabrics.Equals(Fabrics));
                            }
                            if (peise.Count <=0 ){
                            //if (dr[19] is DBNull || Convert.ToInt32(dr[19]) == 0)
                            //{
                                PeiSe ps = new PeiSe() 
                                {
                                    
                                    
                                    PingMing =dr[0].ToString(),
                                    HuoHao = dr[1].ToString(),
                                    GuiGe = dr[2].ToString(),
                                    C61601C1 = dr[3].ToString(),
                                    C61602C1 = dr[4].ToString(),
                                    C61603C1 = dr[5].ToString(),
                                    C61605C1 = dr[6].ToString(),
                                    C61606C1 = dr[7].ToString(),
                                    C61607C1 = dr[8].ToString(),
                                    C61609C1 = dr[9].ToString(),
                                    C61611C1 = dr[10].ToString(),
                                    C61618C1 = dr[11].ToString(),
                                    C61624C1 = dr[12].ToString(),
                                    C61627C1 = dr[13].ToString(),
                                    C61631C1 = dr[14].ToString(),
                                    C61632C1 = dr[15].ToString(),
                                    C61633C1 = dr[16].ToString(),
                                    C61634C1 = dr[17].ToString(),
                                    MianLiaoYanSe = dr[18].ToString(),
                                };
                                if (!ps.PingMing.Equals(string.Empty)) 
                                {
                                    //if (!dr[20].ToString().Equals(string.Empty))
                                    //{
                                    //    ps.Fabrics = dr[20].ToString();
                                    //}
                                    //else
                                    //{
                                        ps.Fabrics = Fabrics;
                                    //}
                                    //if (!dr[21].ToString().Equals(string.Empty))
                                    //{
                                    //    ps.Date = dr[21].ToString();
                                    //}
                                    //else
                                    //{
                                        ps.Date = date;
                                    //}
                                    nep.PeiSe.Add(ps);
                                }
                                
                            }
                            else
                            {
                                int id = Convert.ToInt32(dr[19]);
                                var select = from ks in nep.PeiSe where ks.Id == id select ks;
                                var target = select.FirstOrDefault<PeiSe>();
                                target.Fabrics = dr[20].ToString();
                                target.Date = dr[21].ToString();
                                target.PingMing =dr[0].ToString();
                                target.HuoHao = dr[1].ToString();
                                target.GuiGe = dr[2].ToString();
                                target.C61601C1 = dr[3].ToString();
                                target.C61602C1 = dr[4].ToString();
                                target.C61603C1 = dr[5].ToString();
                                target.C61605C1 = dr[6].ToString();
                                target.C61606C1 = dr[7].ToString();
                                target.C61607C1 = dr[8].ToString();
                                target.C61609C1 = dr[9].ToString();
                                target.C61611C1 = dr[10].ToString();
                                target.C61618C1 = dr[11].ToString();
                                target.C61624C1 = dr[12].ToString();
                                target.C61627C1 = dr[13].ToString();
                                target.C61631C1 = dr[14].ToString();
                                target.C61632C1 = dr[15].ToString();
                                target.C61633C1 = dr[16].ToString();
                                target.C61634C1 = dr[17].ToString();
                                target.MianLiaoYanSe = dr[18].ToString();

                            }
                        }
                        nep.SaveChanges();
                    }
                }
            #endregion

                #region 读取配色数据表EXCEL
                public List<PeiSe> ReaderPeiSe(string fileName) 
                {
                    List<PeiSe> list = new List<PeiSe>();
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
                    {
                        WorkbookPart wbPart = document.WorkbookPart;
                        List<Sheet> sheets = wbPart.Workbook.Descendants<Sheet>().ToList();
                        var versionSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                        WorksheetPart worksheetPart = (WorksheetPart)wbPart.GetPartById(versionSheet.Id);
                        int rowindex = 0;
                        int insertpd = 0;
                        foreach (Row row in worksheetPart.Worksheet.Descendants<Row>())
                        {
                            if (rowindex < 1)
                            {
                                rowindex++;
                                continue;
                            }
                            PeiSe ps = new PeiSe();
                            
                            foreach (Cell cell in row)
                            {
                                string rev = cell.CellReference.Value;
                                if (rev.StartsWith("B") && rev.EndsWith("2")) 
                                {
                                    ps.Fabrics = GetCellValue(wbPart, cell);
                                    //insertpd = 1;
                                }
                                if (rev.StartsWith("B") && rev.EndsWith("3"))
                                {
                                    ps.Date = GetCellValue(wbPart, cell);
                                    insertpd = 1;
                                }
                                if (insertpd == 1)
                                {
                                    if (rowindex >= 5)
                                    {
                                        if (rev.StartsWith("A"))
                                        {
                                            ps.PingMing = GetCellValue(wbPart, cell);
                                        }
                                        if (rev.StartsWith("B"))
                                        {
                                            ps.HuoHao = GetCellValue(wbPart, cell);
                                        }
                                        if (rev.StartsWith("C"))
                                        {
                                            ps.GuiGe = GetCellValue(wbPart, cell);
                                        }
                                        if (rev.StartsWith("D"))
                                        {
                                            ps.C61601C1 = GetCellValue(wbPart, cell);
                                        }
                                        if (rev.StartsWith("E"))
                                        {
                                            ps.C61602C1 = GetCellValue(wbPart, cell);
                                        }
                                        if (rev.StartsWith("F"))
                                        {
                                            ps.C61603C1 = GetCellValue(wbPart, cell);
                                        }
                                        if (rev.StartsWith("G"))
                                        {
                                            ps.C61605C1 = GetCellValue(wbPart, cell);
                                        }
                                        if (rev.StartsWith("H"))
                                        {
                                            ps.C61606C1 = GetCellValue(wbPart, cell);
                                        }
                                        if (rev.StartsWith("I"))
                                        {
                                            ps.C61607C1 = GetCellValue(wbPart, cell);
                                        }
                                        if (rev.StartsWith("J"))
                                        {
                                            ps.C61609C1 = GetCellValue(wbPart, cell);
                                        }
                                        if (rev.StartsWith("K"))
                                        {
                                            ps.C61611C1 = GetCellValue(wbPart, cell);
                                        }
                                        if (rev.StartsWith("L"))
                                        {
                                            ps.C61618C1 = GetCellValue(wbPart, cell);
                                        }
                                        if (rev.StartsWith("M"))
                                        {
                                            ps.C61624C1 = GetCellValue(wbPart, cell);
                                        }
                                        if (rev.StartsWith("N"))
                                        {
                                            ps.C61627C1 = GetCellValue(wbPart, cell);
                                        }
                                        if (rev.StartsWith("O"))
                                        {
                                            ps.C61631C1 = GetCellValue(wbPart, cell);
                                        }
                                        if (rev.StartsWith("P"))
                                        {
                                            ps.C61632C1 = GetCellValue(wbPart, cell);
                                        }
                                        if (rev.StartsWith("Q"))
                                        {
                                            ps.C61633C1 = GetCellValue(wbPart, cell);
                                        }
                                        if (rev.StartsWith("R"))
                                        {
                                            ps.C61634C1 = GetCellValue(wbPart, cell);
                                        }
                                    }
                                    else
                                    {
                                        rowindex++;
                                    }
                                }
                            }
                            list.Add(ps);
                        }
                        return list;
                    }
                }
            #endregion

                #region 删除配色表信息
                public void deletePeiseSession(List<int> id)
                {
                    using (nemanpingEntities3 npe = new nemanpingEntities3())
                    {
                        foreach (int strid in id)
                        {
                            if (strid != null)
                            {
                                if (strid != 0)
                                {
                                    var select = from s in npe.PeiSe where s.Id == strid select s;
                                    foreach (var item in select)
                                    {
                                        npe.PeiSe.Remove(item);
                                    }
                                }
                            }
                        }
                        npe.SaveChanges();
                    }

                }
                #endregion

                #region 删除配色表
                    public void deletps(string caidanhao) 
                {
                    using (nemanpingEntities3 npe = new nemanpingEntities3())
                    {

                        if (caidanhao != null)
                            {
                                var select = from s in npe.PeiSe where s.Fabrics.Equals(caidanhao) select s;
                                foreach (var item in select)
                                {
                                    npe.PeiSe.Remove(item);
                                }
                            }
                        
                        npe.SaveChanges();
                    }
                }
                #endregion
            #endregion

        #region 功能:库存表录入
            #region 查询库存
            public List<KuCun> SelectKC() 
            {
                List<KuCun> list = new List<KuCun>();
                using (nemanpingEntities3 nep = new nemanpingEntities3()) 
                {
                    var select = from kc in nep.KuCun
                                    select kc;
                    list = select.ToList<KuCun>();

                }
                return list;
            }
        #endregion

            #region 提交库存
            public void insertKucun(DataTable dt ) 
            {
                using (nemanpingEntities3 can = new nemanpingEntities3())
                {
                    foreach (DataRow dr in dt.Rows)
                    {

                        if (dr[0] is DBNull || Convert.ToInt32(dr[0]) == 0)
                        {
                            KuCun insets = new KuCun()
                            {
                                PingMing= dr[1].ToString(),
                                HuoHao = dr[2].ToString(),
                                SeHao = dr[3].ToString(),
                                ShuLiang = dr[4].ToString(),
                                GongHuoFang = dr[5].ToString(),
                                CunFangDI = dr[6].ToString()
                            };
                            can.KuCun.Add(insets);

                        }
                        else
                        {
                            int id = Convert.ToInt32(dr[0]);
                            var select = from sc in can.KuCun where sc.Id == id select sc;
                            var target = select.FirstOrDefault<KuCun>();
                            target.PingMing= dr[1].ToString();
                            target.HuoHao = dr[2].ToString();
                            target.SeHao = dr[3].ToString();
                            target.ShuLiang = dr[4].ToString();
                            target.GongHuoFang = dr[5].ToString();
                            target.CunFangDI = dr[6].ToString();
                        }
                    }
                    can.SaveChanges();

                }

            }
        #endregion

            #region 读取EXCEL库存表
            public List<KuCun> readerKucunExcel(string fileName)
            {
                List<KuCun> list = new List<KuCun>();
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
                {
                    WorkbookPart wbPart = document.WorkbookPart;
                    List<Sheet> sheets = wbPart.Workbook.Descendants<Sheet>().ToList();
                    var versionSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                    WorksheetPart worksheetPart = (WorksheetPart)wbPart.GetPartById(versionSheet.Id);
                    int rowindex = 0;
                    foreach (Row row in worksheetPart.Worksheet.Descendants<Row>())
                    {
                        if (rowindex < 1)
                        {
                            rowindex++;
                            continue;
                        }
                        KuCun s = new KuCun();
                        foreach (Cell cell in row)
                        {
                            string rev = cell.CellReference.Value;
                            if (rev.StartsWith("A"))
                            {
                                s.PingMing = GetCellValue(wbPart, cell);
                            }
                            if (rev.StartsWith("B"))
                            {
                                s.HuoHao = GetCellValue(wbPart, cell);
                            }
                            if (rev.StartsWith("C"))
                            {
                                s.SeHao = GetCellValue(wbPart, cell);
                            }
                            if (rev.StartsWith("D"))
                            {
                                s.ShuLiang = GetCellValue(wbPart, cell);
                            }
                            if (rev.StartsWith("E"))
                            {
                                s.GongHuoFang = GetCellValue(wbPart, cell);
                            }
                            if (rev.StartsWith("F"))
                            {
                                s.CunFangDI = GetCellValue(wbPart, cell);
                            }
                        }
                        list.Add(s);
                    }
                    return list;
                }
            }
        #endregion

            #region 删除尺码搭配信息
            public void deleteKucun(List<int> id)
            {
                using (nemanpingEntities3 npe = new nemanpingEntities3())
                {
                    foreach (int strid in id)
                    {
                        if (strid != null)
                        {
                            if (strid != 0)
                            {
                                var select = from s in npe.KuCun where s.Id == strid select s;
                                foreach (var item in select)
                                {
                                    npe.KuCun.Remove(item);
                                }
                            }
                        }
                    }
                    npe.SaveChanges();
                }

            }
            #endregion
    #endregion

        #region 功能 面辅料订购单
            #region 保存
            public void SaveMianFuliaoDingGouDan(DataTable dt) 
            {
                using(nemanpingEntities3 npe = new nemanpingEntities3())
                {
                    foreach (DataRow dr in dt.Rows)
                    {

                        if (dr[0] is DBNull || dr[0].Equals(string.Empty) || Convert.ToInt32(dr[0]) == 0)
                        {
                            MianFuLiaoDingGouDan insets = new MianFuLiaoDingGouDan()
                            {
                                PingMing = dr[1].ToString(),
                                HuoHao = dr[2].ToString(),
                                SeHao = dr[3].ToString(),
                                YanSe = dr[4].ToString(),
                                GuiGe = dr[5].ToString(),
                                DanWei = dr[6].ToString(),
                                DanJia = dr[7].ToString(),
                                ShuLiang = dr[8].ToString(),
                                ZongJinE = dr[9].ToString(),
                                CaiDanHao = dr[10].ToString(),
                                GongFang = dr[11].ToString(),
                                XuFang  = dr[12].ToString(),
                                HeTongHao = dr[13].ToString(),
                                QianYueShiJian = dr[14].ToString(),
                                QianYueDiDan = dr[15].ToString(),
                            };
                            npe.MianFuLiaoDingGouDan.Add(insets);

                        }
                        else
                        {
                            int id = Convert.ToInt32(dr[0]);
                            var select = from sc in npe.MianFuLiaoDingGouDan where sc.Id == id select sc;
                            var target = select.FirstOrDefault<MianFuLiaoDingGouDan>();
                            target.PingMing = dr[1].ToString();
                            target.HuoHao = dr[2].ToString();
                            target.SeHao = dr[3].ToString();
                            target.YanSe = dr[4].ToString();
                            target.GuiGe = dr[5].ToString();
                            target.DanWei = dr[6].ToString();
                            target.DanJia = dr[7].ToString();
                            target.ShuLiang = dr[8].ToString();
                            target.ZongJinE = dr[9].ToString();
                            target.CaiDanHao = dr[10].ToString();
                            target.GongFang = dr[11].ToString();
                            target.XuFang = dr[12].ToString();
                            target.HeTongHao = dr[13].ToString();
                            target.QianYueShiJian = dr[14].ToString();
                            target.QianYueDiDan = dr[15].ToString();
                        }
                    }
                    npe.SaveChanges();


                }
            }
            #endregion

            #region 查询
            public List<MianFuLiaoDingGouDan> SelectMianFuLiao()
            {
                using (nemanpingEntities3 npe = new nemanpingEntities3())
                {
                    List<MianFuLiaoDingGouDan> list = new List<MianFuLiaoDingGouDan>();

                    var select = from cm in npe.MianFuLiaoDingGouDan
                                 select cm;
                    list = select.ToList<MianFuLiaoDingGouDan>();
                    return list;
                }

            }
            #endregion

        #endregion


    }
}
