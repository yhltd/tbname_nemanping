using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using clsBuiness;
using System.Data;
using System.IO;
using NPOI.SS.Util;
namespace logic
{
   public class GongNeng2
    {
        #region 裁单表录入
            #region 生成裁单表
            public List<CaiDan> CreateCaiDan(string style, string chima)
            {
                List<CaiDan> list = new List<CaiDan>();
                using (nemanpingEntities3 npe = new nemanpingEntities3())
                {
                    var selectKuanshi = from kuanshi in npe.KuanShiBiao
                                        where kuanshi.STYLE.Equals(style)
                                        select kuanshi;
                    CaiDan cd = new CaiDan();
                    foreach (var item in selectKuanshi)
                    {
                        cd.DESC = item.DESC;
                        cd.FABRIC = item.FABRIC;
                        cd.STYLE = item.STYLE;
                        cd.Jacket = item.JACKET;
                        cd.Pant = item.PANT;
                        cd.shuoming = item.ShuoMing;
                        list.Add(cd);
                    }

                }
                return list;
            }
            #endregion
            public void InsertCaidan(DataTable dt,string desc,string fabric,string style,string jacket, string pant, string shuoming,string jiagongchang, string caidan,string zhidan ,string jiaohuo,string rn ,string mianliao,string label) 
            {
                using (nemanpingEntities3 can = new nemanpingEntities3())
                {
                    foreach (DataRow dr in dt.Rows)
                    {

                        if (dr[0] is DBNull || Convert.ToInt32(dr[0]) == 0)
                        {
                            CaiDan insets = new CaiDan()
                            {
                               DESC = desc,
                                FABRIC =fabric,
                                STYLE=style,
                                Jacket =jacket,
                                Pant= pant,
                                LABEL = label,
                                shuoming = shuoming,
                                JiaGongchang =jiagongchang,
                                CaiDanHao=caidan,
                                ZhiDanRiqi = zhidan,
                                JiaoHuoRiqi = jiaohuo,
                                RN_NO = rn,
                                MianLiao = mianliao,
                               LOT = dr[1].ToString(),
                               ChimaSTYLE = dr[2].ToString(),
                               ART = dr[3].ToString(),
                               COLOR = dr[4].ToString(),
                               COLORID = dr[5].ToString(),
                               JACKET_PANT = dr[6].ToString(),
                               C34R = dr[7].ToString(),
                               C36R = dr[8].ToString(),
                               C38R = dr[9].ToString(),
                               C40R = dr[10].ToString(),
                               C42R = dr[11].ToString(),
                               C44R = dr[12].ToString(),
                               C46R = dr[13].ToString(),
                               C48R = dr[14].ToString(),
                               C50R = dr[15].ToString(),
                               C52R = dr[16].ToString(),
                               C54R = dr[17].ToString(),
                               C56R = dr[18].ToString(),
                               C58R = dr[19].ToString(),
                               C60R = dr[20].ToString(),
                               C62R = dr[21].ToString(),
                               C36L = dr[22].ToString(),
                               C38L = dr[23].ToString(),
                               C40L = dr[24].ToString(),
                               C42L = dr[25].ToString(),
                               C44L = dr[26].ToString(),
                               C46L = dr[27].ToString(),
                               C48L = dr[28].ToString(),
                               C50L = dr[29].ToString(),
                               C52L = dr[30].ToString(),
                               C54L = dr[31].ToString(),
                               C56L = dr[32].ToString(),
                               C58L = dr[33].ToString(),
                               C60L = dr[34].ToString(),
                               C62L= dr[35].ToString(),
                               C34S = dr[36].ToString(),
                               C36S = dr[37].ToString(),
                               C38S = dr[38].ToString(),
                               C40S = dr[39].ToString(),
                               C42S = dr[40].ToString(),
                               C44S = dr[41].ToString(),
                               C46S = dr[42].ToString(),
                               Sub_Total = dr[43].ToString()
                               
                            };
                            can.CaiDan.Add(insets);

                        }
                        else
                        {
                            int id = Convert.ToInt32(dr[0]);
                            var select = from sc in can.CaiDan where sc.Id == id select sc;
                            var target = select.FirstOrDefault<CaiDan>();
                            target.DESC = desc;
                            target.FABRIC = fabric;
                            target.STYLE = style;
                            target.Jacket = jacket;
                            target.Pant = pant;
                            target.shuoming = shuoming;
                            target.JiaGongchang = jiagongchang;
                            target.CaiDanHao = caidan;
                            target.ZhiDanRiqi = zhidan;
                            target.JiaoHuoRiqi = jiaohuo;
                            target.RN_NO = rn;
                            target.MianLiao = mianliao;
                            target.LOT = dr[1].ToString();
                            target.ChimaSTYLE = dr[2].ToString();
                            target.ART = dr[3].ToString();
                            target.COLOR = dr[4].ToString();
                            target.COLORID = dr[5].ToString();
                            target.JACKET_PANT = dr[6].ToString();
                            target.C34R = dr[7].ToString();
                            target.C36R = dr[8].ToString();
                            target.C38R = dr[9].ToString();
                            target.C40R = dr[10].ToString();
                            target.C42R = dr[11].ToString();
                            target.C44R = dr[12].ToString();
                            target.C46R = dr[13].ToString();
                            target.C48R = dr[14].ToString();
                            target.C50R = dr[15].ToString();
                            target.C52R = dr[16].ToString();
                            target.C54R = dr[17].ToString();
                            target.C56R = dr[18].ToString();
                            target.C58R = dr[19].ToString();
                            target.C60R = dr[20].ToString();
                            target.C62R = dr[21].ToString();
                            target.C36L = dr[22].ToString();
                            target.C38L = dr[23].ToString();
                            target.C40L = dr[24].ToString();
                            target.C42L = dr[25].ToString();
                            target.C44L = dr[26].ToString();
                            target.C46L = dr[27].ToString();
                            target.C48L = dr[28].ToString();
                            target.C50L = dr[29].ToString();
                            target.C52L = dr[30].ToString();
                            target.C54L = dr[31].ToString();
                            target.C56L = dr[32].ToString();
                            target.C58L = dr[33].ToString();
                            target.C60L = dr[34].ToString();
                            target.C62L = dr[35].ToString();
                            target.C34S = dr[36].ToString();
                            target.C36S = dr[37].ToString();
                            target.C38S = dr[38].ToString();
                            target.C40S = dr[39].ToString();
                            target.C42S = dr[40].ToString();
                            target.C44S = dr[41].ToString();
                            target.C46S = dr[42].ToString();
                            target.Sub_Total = dr[43].ToString();
                        }
                    }
                    can.SaveChanges();

                }

            }
            #region 将表格保存至EXCEL
            public void CDEXCEL(DataTable dt,CaiDan cd,string file) 
            {
                string path = Directory.GetCurrentDirectory();
                using (FileStream fs = File.Open(path + "\\Muban\\caidanBiao.xls", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    FileStream patha = File.OpenWrite(file+"\\裁单表-"+cd.CaiDanHao+".xls");
                    HSSFWorkbook wb = new HSSFWorkbook(fs);
                    fs.Close();
                    Sheet st1 = wb.GetSheet("Sheet1");
                    Row r = st1.GetRow(1);
                    Cell descvalue = r.CreateCell(1);
                    descvalue.SetCellValue(cd.DESC);
                    Cell Lablevalue = r.CreateCell(14);
                    Lablevalue.SetCellValue(cd.LABEL);
                    st1.AddMergedRegion(new CellRangeAddress(1, 1, 14, 18));
                    Cell JiaGongChangValue = r.CreateCell(38);
                    JiaGongChangValue.SetCellValue(cd.JiaGongchang);
                    Row r3 = st1.GetRow(2);
                    Cell fabricValue = r3.CreateCell(2);
                    fabricValue.SetCellValue(cd.FABRIC);
                    Cell CaiDanHaoValue = r3.CreateCell(38);
                    CaiDanHaoValue.SetCellValue(cd.CaiDanHao);
                    Row r4 = st1.GetRow(3);
                    Cell StyleVal = r4.CreateCell(1);
                    StyleVal.SetCellValue(cd.STYLE);
                    Cell Jackval = r4.CreateCell(3);
                    Jackval.SetCellValue(cd.Jacket);
                    Cell ZhidanRiqiVal = r4.CreateCell(38);
                    ZhidanRiqiVal.SetCellValue(cd.ZhiDanRiqi);
                    Row r5 = st1.GetRow(4);
                    Cell PantVal = r5.CreateCell(3);
                    PantVal.SetCellValue(cd.Pant);
                    Cell JiaoHuoRiqiVal = r5.CreateCell(38);
                    JiaoHuoRiqiVal.SetCellValue(cd.JiaoHuoRiqi);
                    Row r6 = st1.GetRow(5);
                    Cell shuomingVal = r6.CreateCell(1);
                    shuomingVal.SetCellValue(cd.shuoming);
                    Cell RNVAL = r6.CreateCell(38);
                    RNVAL.SetCellValue(cd.RN_NO);
                    Row R7 = st1.GetRow(6);
                    Cell mianliaoval = R7.CreateCell(38);
                    mianliaoval.SetCellValue(cd.MianLiao);
                    //表内数据
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        Row row = st1.CreateRow(i + 10);
                        for (int j = 1; j < dt.Columns.Count; j++)
                        {
                            Cell cell = row.CreateCell(j-1);
                            cell.SetCellValue(dt.Rows[i][j].ToString());
                        }
                    }
                    wb.Write(patha);//向打开的这个xls文件中写入并保存。  
                    patha.Close();
                }
                InsertCaidan(dt, cd.DESC, cd.FABRIC, cd.STYLE, cd.Jacket, cd.Pant, cd.shuoming, cd.JiaGongchang, cd.CaiDanHao, cd.ZhiDanRiqi, cd.JiaoHuoRiqi, cd.RN_NO, cd.MianLiao,cd.LABEL);

            }
            #endregion

            public List<CaiDan> selectCaiDan(string cdhao) 
            {
                List<CaiDan> list = new List<CaiDan>();
                using (nemanpingEntities3 nep = new nemanpingEntities3()) 
                {
                    if (!cdhao.Equals(string.Empty))
                    {
                        var select = from n in nep.CaiDan
                                     where n.CaiDanHao.Equals(cdhao)
                                     select n;
                        list = select.ToList();
                    }
                    else
                    {
                        var select = from n in nep.CaiDan
                                     select n;
                        list = select.ToList();
                    }
                }
                return list;
            }
        #endregion

            #region 保存配色表为EXCEL
            public void SavePeiSeToExcel(DataTable dt,DataTable dt2, string filePath, string style, string cdno)
            {
                string path = Directory.GetCurrentDirectory();
                if (dt != null && dt.Rows.Count>0)
                {
                    using (FileStream fs = File.Open(path + "\\Muban\\PeiSeBiao.xls", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        FileStream patha = File.OpenWrite(filePath + "\\配色表-" + style + "-" + cdno + ".xls");
                        HSSFWorkbook wb = new HSSFWorkbook(fs);
                        fs.Close();
                        Sheet st1 = wb.GetSheet("Sheet1");
                        Row r1 = st1.GetRow(1);
                        Cell cdNo = r1.CreateCell(1);
                        cdNo.SetCellValue(cdno);//裁单号
                        Row r2 = st1.GetRow(2);
                        Cell STYLE = r2.CreateCell(1);
                        STYLE.SetCellValue(style);//裁单号
                        //表内数据
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            Row row = st1.CreateRow(i + 5);
                            for (int j = 0; j < dt.Columns.Count; j++)
                            {
                                Cell cell = row.CreateCell(j);
                                cell.SetCellValue(dt.Rows[i][j].ToString());
                            }
                        }
                        wb.Write(patha);//向打开的这个xls文件中写入并保存。  
                        patha.Close();
                    }
                }
                if (dt2 != null && dt2.Rows.Count>0) 
                {
                    using (FileStream fs = File.Open(path + "\\Muban\\DanHao.xls", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        FileStream patha = File.OpenWrite(filePath + "\\单耗-" + style + "-" + cdno + ".xls");
                        HSSFWorkbook wb = new HSSFWorkbook(fs);
                        fs.Close();
                        Sheet st1 = wb.GetSheet("Sheet1");
                        //表内数据
                        for (int i = 0; i < dt2.Rows.Count; i++)
                        {
                            Row row = st1.CreateRow(i + 5);
                            for (int j = 0; j < dt2.Columns.Count; j++)
                            {
                                Cell cell = row.CreateCell(j);
                                cell.SetCellValue(dt2.Rows[i][j].ToString());
                            }
                        }
                        wb.Write(patha);//向打开的这个xls文件中写入并保存。  
                        patha.Close();
                    }
                }
            }
            #endregion

    }
}