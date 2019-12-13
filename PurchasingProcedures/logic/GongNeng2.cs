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
        public void InsertCaidan(DataTable dt, string desc, string fabric, string style, string jacket, string pant, string shuoming, string jiagongchang, string caidan, string zhidan, string jiaohuo, string rn, string mianliao, string label)
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
                            FABRIC = fabric,
                            STYLE = style,
                            Jacket = jacket,
                            Pant = pant,
                            LABEL = label,
                            shuoming = shuoming,
                            JiaGongchang = jiagongchang,
                            CaiDanHao = caidan,
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
                            C62L = dr[35].ToString(),
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
        public void rgl2EXCEL(DataTable dt, CaiDan cd, string file)
        {
            string path = Directory.GetCurrentDirectory();
            using (FileStream fs = File.Open(path + "\\Muban\\RGL2.xls", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                FileStream patha = File.OpenWrite(file + "\\裁单表-" + cd.CaiDanHao + ".xls");
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
                        Cell cell = row.CreateCell(j - 1);
                        cell.SetCellValue(dt.Rows[i][j].ToString());
                    }
                }
                wb.Write(patha);//向打开的这个xls文件中写入并保存。  
                patha.Close();
            }
            // InsertCaidan(dt, cd.DESC, cd.FABRIC, cd.STYLE, cd.Jacket, cd.Pant, cd.shuoming, cd.InsertCaidan(dt, cd.DESC, cd.FABRIC, cd.STYLE, cd.Jacket, cd.Pant, cd.shuoming, cd.JiaGongchang, cd.CaiDanHao, cd.ZhiDanRiqi, cd.JiaoHuoRiqi, cd.RN_NO, cd.MianLiao, cd.LABEL));
            InsertCaidan_RGL2(dt, cd.DESC, cd.FABRIC, cd.STYLE, cd.Jacket, cd.Pant, cd.shuoming, cd.JiaGongchang, cd.CaiDanHao, cd.ZhiDanRiqi, cd.JiaoHuoRiqi, cd.RN_NO, cd.MianLiao, cd.LABEL);

        }

        public void CDEXCEL(DataTable dt, CaiDan cd, string file)
        {
            string path = Directory.GetCurrentDirectory();
            using (FileStream fs = File.Open(path + "\\Muban\\caidanBiao.xls", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                FileStream patha = File.OpenWrite(file + "\\裁单表-" + cd.CaiDanHao + ".xls");
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
                        Cell cell = row.CreateCell(j - 1);
                        cell.SetCellValue(dt.Rows[i][j].ToString());
                    }
                }
                wb.Write(patha);//向打开的这个xls文件中写入并保存。  
                patha.Close();
            }
            InsertCaidan(dt, cd.DESC, cd.FABRIC, cd.STYLE, cd.Jacket, cd.Pant, cd.shuoming, cd.JiaGongchang, cd.CaiDanHao, cd.ZhiDanRiqi, cd.JiaoHuoRiqi, cd.RN_NO, cd.MianLiao, cd.LABEL);

        }
        public void CDEXCELRGLJ(DataTable dt, CaiDan_RGLJ cd, string file)
        {
            string path = Directory.GetCurrentDirectory();
            using (FileStream fs = File.Open(path + "\\Muban\\RGLJ.xls", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                FileStream patha = File.OpenWrite(file + "\\裁单表-" + cd.CaiDanHao + ".xls");
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
                        Cell cell = row.CreateCell(j - 1);
                        cell.SetCellValue(dt.Rows[i][j].ToString());
                    }
                }
                wb.Write(patha);//向打开的这个xls文件中写入并保存。  
                patha.Close();
            }
            InsertCaidan_RGLJ(dt, cd.DESC, cd.FABRIC, cd.STYLE, cd.Jacket, cd.Pant, cd.shuoming, cd.JiaGongchang, cd.CaiDanHao, cd.ZhiDanRiqi, cd.JiaoHuoRiqi, cd.RN_NO, cd.MianLiao, cd.LABEL);

        }
        public void CDEXCELD_PANT(DataTable dt, CaiDan_D_PANT cd, string file)
        {
            string path = Directory.GetCurrentDirectory();
            using (FileStream fs = File.Open(path + "\\Muban\\D_PANT.xls", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                FileStream patha = File.OpenWrite(file + "\\裁单表-" + cd.CaiDanHao + ".xls");
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
                    Row row = st1.CreateRow(i + 11);
                    for (int j = 1; j < dt.Columns.Count; j++)
                    {
                        Cell cell = row.CreateCell(j - 1);
                        cell.SetCellValue(dt.Rows[i][j - 1].ToString());
                    }
                }
                wb.Write(patha);//向打开的这个xls文件中写入并保存。  
                patha.Close();
            }
            InsertCaidanC_PANT(dt, cd.DESC, cd.FABRIC, cd.STYLE, cd.Jacket, cd.Pant, cd.shuoming, cd.JiaGongchang, cd.CaiDanHao, cd.ZhiDanRiqi, cd.JiaoHuoRiqi, cd.RN_NO, cd.MianLiao, cd.LABEL);

        }


        public void CDEXCELC_PANT(DataTable dt, CaiDan_C_PANT cd, string file)
        {
            string path = Directory.GetCurrentDirectory();
            using (FileStream fs = File.Open(path + "\\Muban\\C_PANT.xls", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                FileStream patha = File.OpenWrite(file + "\\裁单表-" + cd.CaiDanHao + ".xls");
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
                        Cell cell = row.CreateCell(j - 1);
                        cell.SetCellValue(dt.Rows[i][j - 1].ToString());
                    }
                }
                wb.Write(patha);//向打开的这个xls文件中写入并保存。  
                patha.Close();
            }
            InsertCaidanC_PANT(dt, cd.DESC, cd.FABRIC, cd.STYLE, cd.Jacket, cd.Pant, cd.shuoming, cd.JiaGongchang, cd.CaiDanHao, cd.ZhiDanRiqi, cd.JiaoHuoRiqi, cd.RN_NO, cd.MianLiao, cd.LABEL);

        }
        #endregion
        public void CDEXCELSLIM(DataTable dt, CaiDan_SLIM cd, string file)
        {
            string path = Directory.GetCurrentDirectory();
            using (FileStream fs = File.Open(path + "\\Muban\\SLIM.xls", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                FileStream patha = File.OpenWrite(file + "\\裁单表-" + cd.CaiDanHao + ".xls");
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
                        Cell cell = row.CreateCell(j - 1);
                        cell.SetCellValue(dt.Rows[i][j].ToString());
                    }
                }
                wb.Write(patha);//向打开的这个xls文件中写入并保存。  
                patha.Close();
            }
            InsertCaidan_SlIM(dt, cd.DESC, cd.FABRIC, cd.STYLE, cd.Jacket, cd.Pant, cd.shuoming, cd.JiaGongchang, cd.CaiDanHao, cd.ZhiDanRiqi, cd.JiaoHuoRiqi, cd.RN_NO, cd.MianLiao, cd.LABEL);

        }
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
        public List<CaiDan_RGL2> selectCaiDanRGL2(string cdhao)
        {
            List<CaiDan_RGL2> list = new List<CaiDan_RGL2>();
            using (nemanpingEntities3 nep = new nemanpingEntities3())
            {
                if (!cdhao.Equals(string.Empty))
                {
                    var select = from n in nep.CaiDan_RGL2
                                 where n.CaiDanHao.Equals(cdhao)
                                 select n;
                    list = select.ToList();
                }
                else
                {
                    var select = from n in nep.CaiDan_RGL2
                                 select n;
                    list = select.ToList();
                }
            }
            return list;
        }
        public List<CaiDan_SLIM> selectCaiDanSLIM(string cdhao)
        {
            List<CaiDan_SLIM> list = new List<CaiDan_SLIM>();
            using (nemanpingEntities3 nep = new nemanpingEntities3())
            {
                if (!cdhao.Equals(string.Empty))
                {
                    var select = from n in nep.CaiDan_SLIM
                                 where n.CaiDanHao.Equals(cdhao)
                                 select n;
                    list = select.ToList();
                }
                else
                {
                    var select = from n in nep.CaiDan_SLIM
                                 select n;
                    list = select.ToList();
                }
            }
            return list;
        }
        public List<CaiDan_RGLJ> selectCaiDanRGLJ(string cdhao)
        {
            List<CaiDan_RGLJ> list = new List<CaiDan_RGLJ>();
            using (nemanpingEntities3 nep = new nemanpingEntities3())
            {
                if (!cdhao.Equals(string.Empty))
                {
                    var select = from n in nep.CaiDan_RGLJ
                                 where n.CaiDanHao.Equals(cdhao)
                                 select n;
                    list = select.ToList();
                }
                else
                {
                    var select = from n in nep.CaiDan_RGLJ
                                 select n;
                    list = select.ToList();
                }
            }
            return list;
        }
        public List<CaiDan_D_PANT> selectCaiDanD_PANT(string cdhao)
        {
            List<CaiDan_D_PANT> list = new List<CaiDan_D_PANT>();
            using (nemanpingEntities3 nep = new nemanpingEntities3())
            {
                if (!cdhao.Equals(string.Empty))
                {
                    var select = from n in nep.CaiDan_D_PANT
                                 where n.CaiDanHao.Equals(cdhao)
                                 select n;
                    list = select.ToList();
                }
                else
                {
                    var select = from n in nep.CaiDan_D_PANT
                                 select n;
                    list = select.ToList();
                }
            }
            return list;
        }
        public List<CaiDan_C_PANT> selectCaiDanC_PANT(string cdhao)
        {
            List<CaiDan_C_PANT> list = new List<CaiDan_C_PANT>();
            using (nemanpingEntities3 nep = new nemanpingEntities3())
            {
                if (!cdhao.Equals(string.Empty))
                {
                    var select = from n in nep.CaiDan_C_PANT
                                 where n.CaiDanHao.Equals(cdhao)
                                 select n;
                    list = select.ToList();
                }
                else
                {
                    var select = from n in nep.CaiDan_C_PANT
                                 select n;
                    list = select.ToList();
                }
            }
            return list;
        }

        public List<CaiDan_RGL2> selectCaiDanRGLall(string cdhao)
        {
            List<CaiDan_RGL2> list = new List<CaiDan_RGL2>();
            using (nemanpingEntities3 nep = new nemanpingEntities3())
            {
                if (!cdhao.Equals(string.Empty))
                {
                    var select = from n in nep.CaiDan_RGL2
                                 where n.CaiDanHao.Equals(cdhao)
                                 select n;
                    list = select.ToList();
                }
                else
                {
                    var select = from n in nep.CaiDan_RGL2
                                 select n;
                    list = select.ToList();
                }
            }
            return list;
        }
        #endregion

        #region 保存配色表为EXCEL
        public void SavePeiSeToExcel(DataTable dt, DataTable dt2, DataTable dt3, string filePath, string style, string cdno)
        {
            string path = Directory.GetCurrentDirectory();
            if (dt != null && dt.Rows.Count > 0)
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
            if (dt2 != null && dt2.Rows.Count > 0)
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
            if (dt3 != null && dt3.Rows.Count > 0)
            {
                using (FileStream fs = File.Open(path + "\\Muban\\HeDingChengBen.xls", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    FileStream patha = File.OpenWrite(filePath + "\\核定成本-" + style + "-" + cdno + ".xls");
                    HSSFWorkbook wb = new HSSFWorkbook(fs);
                    fs.Close();
                    Sheet st1 = wb.GetSheet("Sheet1");
                    //表内数据
                    for (int i = 0; i < dt3.Rows.Count; i++)
                    {
                        Row row = st1.CreateRow(i + 4);
                        for (int j = 0; j < dt3.Columns.Count; j++)
                        {

                            Cell cell = row.CreateCell(j);
                            cell.SetCellValue(dt3.Rows[i][j].ToString());
                        }
                    }
                    wb.Write(patha);//向打开的这个xls文件中写入并保存。  
                    patha.Close();
                }
            }
        }
        #endregion

        public void InsertCaidan_SlIM(DataTable dt, string desc, string fabric, string style, string jacket, string pant, string shuoming, string jiagongchang, string caidan, string zhidan, string jiaohuo, string rn, string mianliao, string label)
        {
            using (nemanpingEntities3 can = new nemanpingEntities3())
            {
                foreach (DataRow dr in dt.Rows)
                {

                    if (dr[0] is DBNull || Convert.ToInt32(dr[0]) == 0)
                    {
                        CaiDan_SLIM insets = new CaiDan_SLIM()
                        {
                            DESC = desc,
                            FABRIC = fabric,
                            STYLE = style,
                            Jacket = jacket,
                            Pant = pant,
                            LABEL = label,
                            shuoming = shuoming,
                            JiaGongchang = jiagongchang,
                            CaiDanHao = caidan,
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
                            C36L = dr[15].ToString(),
                            C38L = dr[16].ToString(),
                            C40L = dr[17].ToString(),
                            C42L = dr[18].ToString(),
                            C44L = dr[19].ToString(),
                            C46L = dr[20].ToString(),
                            C48L = dr[21].ToString(),
                            C34S = dr[22].ToString(),
                            C36S = dr[23].ToString(),
                            C38S = dr[24].ToString(),
                            C40S = dr[25].ToString(),
                            C42S = dr[26].ToString(),
                            C44S = dr[27].ToString(),
                            C46S = dr[28].ToString(),
                            Sub_Total = dr[29].ToString()

                        };
                        can.CaiDan_SLIM.Add(insets);

                    }
                    else
                    {
                        int id = Convert.ToInt32(dr[0]);
                        var select = from sc in can.CaiDan_SLIM where sc.id == id select sc;
                        var target = select.FirstOrDefault<CaiDan_SLIM>();
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
                        target.C36L = dr[15].ToString();
                        target.C38L = dr[16].ToString();
                        target.C40L = dr[17].ToString();
                        target.C42L = dr[18].ToString();
                        target.C44L = dr[19].ToString();
                        target.C46L = dr[20].ToString();
                        target.C48L = dr[21].ToString();
                        target.C34S = dr[22].ToString();
                        target.C36S = dr[23].ToString();
                        target.C38S = dr[24].ToString();
                        target.C40S = dr[25].ToString();
                        target.C42S = dr[26].ToString();
                        target.C44S = dr[27].ToString();
                        target.C46S = dr[28].ToString();

                        target.Sub_Total = dr[29].ToString();
                    }
                }
                can.SaveChanges();

            }
        }
        public void InsertCaidanC_PANT(DataTable dt, string desc, string fabric, string style, string jacket, string pant, string shuoming, string jiagongchang, string caidan, string zhidan, string jiaohuo, string rn, string mianliao, string label)
        {
            using (nemanpingEntities3 can = new nemanpingEntities3())
            {
                foreach (DataRow dr in dt.Rows)
                {

                    if (dr[42] is DBNull || Convert.ToInt32(dr[42]) == 0)
                    {
                        CaiDan_C_PANT insets = new CaiDan_C_PANT()
                        {
                            DESC = desc,
                            FABRIC = fabric,
                            STYLE = style,
                            Jacket = jacket,
                            Pant = pant,
                            LABEL = label,
                            shuoming = shuoming,
                            JiaGongchang = jiagongchang,
                            CaiDanHao = caidan,
                            ZhiDanRiqi = zhidan,
                            JiaoHuoRiqi = jiaohuo,
                            RN_NO = rn,
                            MianLiao = mianliao,
                            LOT = dr[0].ToString(),
                            ChimaSTYLE = dr[1].ToString(),
                            ART = dr[2].ToString(),
                            COLOR = dr[3].ToString(),
                            COLORID = dr[4].ToString(),
                            JACKET_PANT = dr[5].ToString(),
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
                        can.CaiDan_C_PANT.Add(insets);

                    }
                    else
                    {
                        int id = Convert.ToInt32(dr[43]);
                        var select = from sc in can.CaiDan_C_PANT where sc.id == id select sc;
                        var target = select.FirstOrDefault<CaiDan_C_PANT>();
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
                        target.LOT = dr[0].ToString();
                        target.ChimaSTYLE = dr[1].ToString();
                        target.ART = dr[2].ToString();
                        target.COLOR = dr[3].ToString();
                        target.COLORID = dr[4].ToString();
                        target.JACKET_PANT = dr[5].ToString();
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
                can.SaveChanges();

            }
        }
        public void InsertCaidanD_PANT(DataTable dt, string desc, string fabric, string style, string jacket, string pant, string shuoming, string jiagongchang, string caidan, string zhidan, string jiaohuo, string rn, string mianliao, string label)
        {
            using (nemanpingEntities3 can = new nemanpingEntities3())
            {
                foreach (DataRow dr in dt.Rows)
                {

                    if (dr[42] is DBNull || Convert.ToInt32(dr[42]) == 0)
                    {
                        CaiDan_D_PANT insets = new CaiDan_D_PANT()
                        {
                            DESC = desc,
                            FABRIC = fabric,
                            STYLE = style,
                            Jacket = jacket,
                            Pant = pant,
                            LABEL = label,
                            shuoming = shuoming,
                            JiaGongchang = jiagongchang,
                            CaiDanHao = caidan,
                            ZhiDanRiqi = zhidan,
                            JiaoHuoRiqi = jiaohuo,
                            RN_NO = rn,
                            MianLiao = mianliao,
                            LOT = dr[0].ToString(),
                            ChimaSTYLE = dr[1].ToString(),
                            ART = dr[2].ToString(),
                            COLOR = dr[3].ToString(),
                            COLORID = dr[4].ToString(),
                            JACKET_PANT = dr[5].ToString(),
                            C30W_R_30L = dr[6].ToString(),
                            C30W_L_32L = dr[7].ToString(),
                            C32W_R_30L = dr[8].ToString(),
                            C32W_L_32L = dr[9].ToString(),
                            C34W_S_28L = dr[10].ToString(),
                            C34W_S_29L = dr[11].ToString(),
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
                            Sub_Total = dr[41].ToString()

                        };
                        can.CaiDan_D_PANT.Add(insets);

                    }
                    else
                    {
                        int id = Convert.ToInt32(dr[43]);
                        var select = from sc in can.CaiDan_D_PANT where sc.id == id select sc;
                        var target = select.FirstOrDefault<CaiDan_D_PANT>();
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
                        target.LOT = dr[0].ToString();
                        target.ChimaSTYLE = dr[1].ToString();
                        target.ART = dr[2].ToString();
                        target.COLOR = dr[3].ToString();
                        target.COLORID = dr[4].ToString();
                        target.JACKET_PANT = dr[5].ToString();
                        target.C30W_R_30L = dr[6].ToString();
                        target.C30W_L_32L = dr[7].ToString();
                        target.C32W_R_30L = dr[8].ToString();
                        target.C32W_L_32L = dr[9].ToString();
                        target.C34W_S_28L = dr[10].ToString();
                        target.C34W_S_29L = dr[11].ToString();
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
                can.SaveChanges();

            }
        }
        public void InsertCaidan_RGLJ(DataTable dt, string desc, string fabric, string style, string jacket, string pant, string shuoming, string jiagongchang, string caidan, string zhidan, string jiaohuo, string rn, string mianliao, string label)
        {
            using (nemanpingEntities3 can = new nemanpingEntities3())
            {
                foreach (DataRow dr in dt.Rows)
                {

                    if (dr[0] is DBNull || Convert.ToInt32(dr[0]) == 0)
                    {
                        CaiDan_RGLJ insets = new CaiDan_RGLJ()
                        {
                            DESC = desc,
                            FABRIC = fabric,
                            STYLE = style,
                            Jacket = jacket,
                            Pant = pant,
                            LABEL = label,
                            shuoming = shuoming,
                            JiaGongchang = jiagongchang,
                            CaiDanHao = caidan,
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
                            C62L = dr[35].ToString(),
                            C34S = dr[36].ToString(),
                            C36S = dr[37].ToString(),
                            C38S = dr[38].ToString(),
                            C40S = dr[39].ToString(),
                            C42S = dr[40].ToString(),
                            C44S = dr[41].ToString(),
                            C46S = dr[42].ToString(),
                            Sub_Total = dr[43].ToString()

                        };
                        can.CaiDan_RGLJ.Add(insets);

                    }
                    else
                    {
                        int id = Convert.ToInt32(dr[0]);
                        var select = from sc in can.CaiDan_RGLJ where sc.id == id select sc;
                        var target = select.FirstOrDefault<CaiDan_RGLJ>();
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
        #region yao

        public void InsertCaidan_RGL2(DataTable dt, string desc, string fabric, string style, string jacket, string pant, string shuoming, string jiagongchang, string caidan, string zhidan, string jiaohuo, string rn, string mianliao, string label)
        {
            using (nemanpingEntities3 can = new nemanpingEntities3())
            {
                foreach (DataRow dr in dt.Rows)
                {

                    if (dr[0] is DBNull || Convert.ToInt32(dr[0]) == 0)
                    {
                        CaiDan_RGL2 insets = new CaiDan_RGL2()
                        {
                            DESC = desc,
                            FABRIC = fabric,
                            STYLE = style,
                            Jacket = jacket,
                            Pant = pant,
                            LABEL = label,
                            shuoming = shuoming,
                            JiaGongchang = jiagongchang,
                            CaiDanHao = caidan,
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
                            C62L = dr[35].ToString(),
                            C34S = dr[36].ToString(),
                            C36S = dr[37].ToString(),
                            C38S = dr[38].ToString(),
                            C40S = dr[39].ToString(),
                            C42S = dr[40].ToString(),
                            C44S = dr[41].ToString(),
                            C46S = dr[42].ToString(),
                            Sub_Total = dr[43].ToString()

                        };
                        can.CaiDan_RGL2.Add(insets);

                    }
                    else
                    {
                        int id = Convert.ToInt32(dr[0]);
                        var select = from sc in can.CaiDan_RGL2 where sc.Id == id select sc;
                        var target = select.FirstOrDefault<CaiDan_RGL2>();
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

        #endregion
    }
}