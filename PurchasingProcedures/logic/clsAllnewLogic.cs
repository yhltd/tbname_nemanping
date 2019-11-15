using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using clsBuiness;
using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
namespace logic
{
    public class clsAllnewLogic
    {
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

                        if (dr[0] is DBNull || Convert.ToInt32(dr[0]) == 0)
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
                            int id = Convert.ToInt32(dr[0]);
                            var select = from sc in can.Sehao where sc.Id == id select sc;
                            var target = select.FirstOrDefault<Sehao>();
                            target.Name = dr[1].ToString();
                            target.SeHao1 = dr[2].ToString();

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
            #endregion

            #region 添加修改 尺码搭配表
                public void InsertChima(DataTable dt,string biaogeName) 
                {
                    using (nemanpingEntities3 npe = new nemanpingEntities3()) 
                    {
                        foreach (DataRow dr in dt.Rows) 
                        {
                            if (dr[0] is DBNull || Convert.ToInt32(dr[0]) == 0)
                            {
                                ChiMa_Dapeibiao cd = new ChiMa_Dapeibiao()
                                {
                                    BiaoName = biaogeName,
                                    LOT__面料 = dr[1].ToString(),
                                    STYLE_款式 = dr[2].ToString(),
                                    ART_货号 = dr[3].ToString(),
                                    COLOR_颜色 = dr[4].ToString(),
                                    COLOR__颜色编号 = dr[5].ToString(),
                                    JACKET_上衣_PANT_裤子 = dr[6].ToString(),
                                    C34R_28 = dr[7].ToString(),
                                    C36R_30 = dr[8].ToString(),
                                    C38R_32 = dr[9].ToString(),
                                    C40R___34 = dr[10].ToString(),
                                    C42R_36 = dr[11].ToString(),
                                    C44R_38 = dr[12].ToString(),
                                    C46R_40 = dr[13].ToString(),
                                    C48R_42 = dr[14].ToString(),
                                    C50R_44 = dr[15].ToString(),
                                    C52R_46 = dr[16].ToString(),
                                    C54R_48 = dr[17].ToString(),
                                    C56R_50 = dr[18].ToString(),
                                    C58R_52 = dr[19].ToString(),
                                    C60R_54 = dr[20].ToString(),
                                    C62R_56 = dr[21].ToString(),
                                    C36L_30 = dr[22].ToString(),
                                    C38L_32 = dr[23].ToString(),
                                    C40L_34 = dr[24].ToString(),
                                    C42L_36 = dr[25].ToString(),
                                    C44L_38 = dr[26].ToString(),
                                    C46L_40 = dr[27].ToString(),
                                    C48L_42 = dr[28].ToString(),
                                    C50L_44 = dr[29].ToString(),
                                    C52L_46 = dr[30].ToString(),
                                    C54L_48 = dr[31].ToString(),
                                    C56L_50 = dr[32].ToString(),
                                    C58L_52 = dr[33].ToString(),
                                    C60L_54 = dr[34].ToString(),
                                    C62L_56 = dr[35].ToString(),
                                    C34S_28 = dr[36].ToString(),
                                    C36S_30 = dr[37].ToString(),
                                    C38S_32 = dr[38].ToString(),
                                    C40S_34 = dr[39].ToString(),
                                    C42S_36 = dr[40].ToString(),
                                    C44S_38 = dr[41].ToString(),
                                    C46S_40 = dr[42].ToString(),
                                    DingdanHeji = dr[43].ToString()
                                };
                                npe.ChiMa_Dapeibiao.Add(cd);
                            }
                            else 
                            {
                                int id = Convert.ToInt32(dr[0].ToString());
                                var select = from cd in npe.ChiMa_Dapeibiao
                                             where cd.id ==id
                                             select cd;
                                var target = select.FirstOrDefault<ChiMa_Dapeibiao>();
                                target.LOT__面料 = dr[1].ToString();
                                target.STYLE_款式 = dr[2].ToString();
                                target.ART_货号 = dr[3].ToString();
                                target.COLOR_颜色 = dr[4].ToString();
                                target.COLOR__颜色编号 = dr[5].ToString();
                                target.JACKET_上衣_PANT_裤子 = dr[6].ToString();
                                target.C34R_28 = dr[7].ToString();
                                target. C36R_30 = dr[8].ToString();
                                target.C38R_32 = dr[9].ToString();
                                target.C40R___34 = dr[10].ToString();
                                target.C42R_36 = dr[11].ToString();
                                target.C44R_38 = dr[12].ToString();
                                target.C46R_40 = dr[13].ToString();
                                target.C48R_42 = dr[14].ToString();
                                target.C50R_44 = dr[15].ToString();
                                target.C52R_46 = dr[16].ToString();
                                target.C54R_48 = dr[17].ToString();
                                target.C56R_50 = dr[18].ToString();
                                target.C58R_52 = dr[19].ToString();
                                target.C60R_54 = dr[20].ToString();
                                target.C62R_56 = dr[21].ToString();
                                target.C36L_30 = dr[22].ToString();
                                target.C38L_32 = dr[23].ToString();
                                target.C40L_34 = dr[24].ToString();
                                target.C42L_36 = dr[25].ToString();
                                target.C44L_38 = dr[26].ToString();
                                target.C46L_40 = dr[27].ToString();
                                target.C48L_42 = dr[28].ToString();
                                target.C50L_44 = dr[29].ToString();
                                target.C52L_46 = dr[30].ToString();
                                target.C54L_48 = dr[31].ToString();
                                target.C56L_50 = dr[32].ToString();
                                target.C58L_52 = dr[33].ToString();
                                target.C60L_54 = dr[34].ToString();
                                target.C62L_56 = dr[35].ToString();
                                target.C34S_28 = dr[36].ToString();
                                target.C36S_30 = dr[37].ToString();
                                target.C38S_32 = dr[38].ToString();
                                target.C40S_34 = dr[39].ToString();
                                target.C42S_36 = dr[40].ToString();
                                target.C44S_38 = dr[41].ToString();
                                target.C46S_40 = dr[42].ToString();
                                target.DingdanHeji = dr[43].ToString();  
                            }
                        }
                        npe.SaveChanges();
                    }
                }
            #endregion

            #region 读取 EXCEL尺码搭配表
                public List<ChiMa_Dapeibiao> ReaderChiMaDapei(string fileName) 
                {
                    List<ChiMa_Dapeibiao> list = new List<ChiMa_Dapeibiao>();
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
                    {
                        WorkbookPart wbPart = document.WorkbookPart;
                        List<Sheet> sheets = wbPart.Workbook.Descendants<Sheet>().ToList();
                        var versionSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                        WorksheetPart worksheetPart = (WorksheetPart)wbPart.GetPartById(versionSheet.Id);
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
                        foreach (DataRow dr in dt.Rows)
                        {

                            if (dr[0] is DBNull || Convert.ToInt32(dr[0]) == 0)
                            {
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
                                var select = from sc in nep.DanHao where sc.Id == id select sc;
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
                        foreach (DataRow dr in dt.Rows)
                        {

                            if (dr[19] is DBNull || Convert.ToInt32(dr[19]) == 0)
                            {
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
                                    if (!dr[20].ToString().Equals(string.Empty))
                                    {
                                        ps.Fabrics = dr[20].ToString();
                                    }
                                    else
                                    {
                                        ps.Fabrics = Fabrics;
                                    }
                                    if (!dr[21].ToString().Equals(string.Empty))
                                    {
                                        ps.Date = dr[21].ToString();
                                    }
                                    else
                                    {
                                        ps.Date = date;
                                    }
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
                            if (rowindex < 2)
                            {
                                rowindex++;
                                continue;
                            }
                            PeiSe ps = new PeiSe();
                            
                            foreach (Cell cell in row)
                            {
                                string rev = cell.CellReference.Value;
                                if (rev.StartsWith("C") && rowindex ==2 ) 
                                {
                                    ps.Fabrics = GetCellValue(wbPart, cell);
                                    //insertpd = 1;
                                }
                                if (rev.StartsWith("P") && rowindex == 2)
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



    }
}
