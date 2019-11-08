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
        public List<ChiMa_Dapeibiao> SelectChiMaDapei() 
        {
            using(nemanpingEntities3 npe = new nemanpingEntities3())
            {
                List<ChiMa_Dapeibiao> list = new List<ChiMa_Dapeibiao>();

                var select = from cm in npe.ChiMa_Dapeibiao
                             select cm;
                list = select.ToList<ChiMa_Dapeibiao>();
                return list;
            }
            
        }
            #endregion

            #region 添加修改 尺码搭配表
                public void InsertChima(DataTable dt) 
                {
                    using (nemanpingEntities3 npe = new nemanpingEntities3()) 
                    {
                        foreach (DataRow dr in dt.Rows) 
                        {
                            if (dr[0] is DBNull || Convert.ToInt32(dr[0]) == 0)
                            {
                                ChiMa_Dapeibiao cd = new ChiMa_Dapeibiao()
                                {
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
        #endregion

        #region 功能:单耗表录入
            #region 查询单耗表
                public List<DanHao> SelectDanHao() 
                {
                    using (nemanpingEntities3 npe = new nemanpingEntities3())
                    {
                        List<DanHao> list = new List<DanHao>();

                        var select = from cm in npe.DanHao
                                     select cm;
                        list = select.ToList<DanHao>();
                        return list;
                    }
                }
            #endregion
        #endregion
        #region 功能:配色表录入
            #region 查询配色表
                public List<PeiSe> selectPeise(int strwhere)
                {
                    using (nemanpingEntities3 can = new nemanpingEntities3())
                    {
                        List<PeiSe> sehao = new List<PeiSe>();
                        if (strwhere ==0)
                        {
                            var select = from s in can.PeiSe select s;
                            sehao = select.ToList<PeiSe>();
                        }
                        else 
                        {
                            var select = from s in can.PeiSe where s.Id == strwhere select s;
                            sehao = select.ToList<PeiSe>();
                        }
                       

                        return sehao;
                    }
                }
            #endregion
        #endregion
    }
}
