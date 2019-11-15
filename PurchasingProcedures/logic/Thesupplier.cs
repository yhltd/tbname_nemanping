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
    public class Define1
    {
        public bool Login(string name, string pwd)
        {
            using (nemanpingEntities3 can = new nemanpingEntities3())
            {
                var user = can.UserTable.First(u => u.Name.Equals(name) && u.Pwd.Equals(pwd));
                return true;
            }

        }
        #region 查询供货方信息
        public List<GongHuoFang> selectGongHuoFang()
        {
            using (nemanpingEntities3 can = new nemanpingEntities3())
            {
                List<GongHuoFang> GongHuoFang = new List<GongHuoFang>();
                var select = from s in can.GongHuoFang select s;
                foreach (var item in select)
                {
                    GongHuoFang s = new GongHuoFang();
                    s.Id = item.Id;
                    s.PingMing = item.PingMing;
                    s.HuoHao = item.HuoHao;
                    s.Guige = item.Guige;
                    s.SeHao = item.SeHao;
                    s.Yanse = item.Yanse;
                    s.DanJia = item.DanJia;
                    s.GongHuoFangA = item.GongHuoFangA;
                    s.GongHuoFangB = item.GongHuoFangB;
                    s.BeiZhu = item.BeiZhu;
                    GongHuoFang.Add(s);
                }
                return GongHuoFang;
            }
        }

        #endregion

        #region 添加工厂信息
        public void insertGongHuoFang(DataTable dt)
        {
            using (nemanpingEntities3 can = new nemanpingEntities3())
            {
                foreach (DataRow dr in dt.Rows)
                {

                    if (dr[0] is DBNull || Convert.ToInt32(dr[0]) == 0)
                    {
                        GongHuoFang insets = new GongHuoFang()
                        {
                            PingMing = dr[1].ToString(),
                            HuoHao = dr[2].ToString(),
                            Guige = dr[3].ToString(),
                            SeHao = dr[4].ToString(),
                            Yanse = dr[5].ToString(),
                            DanJia = dr[6].ToString(),
                            GongHuoFangA = dr[7].ToString(),
                            GongHuoFangB = dr[8].ToString(),
                            BeiZhu = dr[9].ToString()

                        };
                        can.GongHuoFang.Add(insets);

                    }
                    else
                    {
                        int id = Convert.ToInt32(dr[0]);
                        var select = from sc in can.GongHuoFang where sc.Id == id select sc;
                        var target = select.FirstOrDefault<GongHuoFang>();
                        target.PingMing = dr[1].ToString();
                        target.HuoHao = dr[2].ToString();
                        target.Guige = dr[3].ToString();
                        target.SeHao = dr[4].ToString();
                        target.Yanse = dr[5].ToString();
                        target.DanJia = dr[6].ToString();
                        target.GongHuoFangA = dr[7].ToString();
                        target.GongHuoFangB = dr[8].ToString();
                        target.BeiZhu = dr[9].ToString();

                    }
                }
                can.SaveChanges();

            }
        }
        #endregion

        #region 读取色号表
        public List<GongHuoFang> readerGongHuoFangExcel(string fileName)
        {
            List<GongHuoFang> list = new List<GongHuoFang>();
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
                    GongHuoFang s = new GongHuoFang();
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
                            s.Guige = GetCellValue(wbPart, cell);
                        }
                        if (rev.StartsWith("D"))
                        {
                            s.SeHao = GetCellValue(wbPart, cell);
                        }
                        if (rev.StartsWith("E"))
                        {
                            s.Yanse = GetCellValue(wbPart, cell);
                        }
                        if (rev.StartsWith("F"))
                        {
                            s.DanJia = GetCellValue(wbPart, cell);
                        }
                        if (rev.StartsWith("G"))
                        {
                            s.GongHuoFangA = GetCellValue(wbPart, cell);
                        }
                        if (rev.StartsWith("H"))
                        {
                            s.GongHuoFangB = GetCellValue(wbPart, cell);

                        }
                        if (rev.StartsWith("I"))
                        {
                            s.BeiZhu = GetCellValue(wbPart, cell);
                        }

                    }
                    list.Add(s);
                }
                return list;
            }
        }
        #endregion

        #region 删除供货方信息
        public void deleteGongHuoFang(List<int> id)
        {
            using (nemanpingEntities3 npe = new nemanpingEntities3())
            {
                foreach (int strid in id)
                {
                    if (strid != null)
                    {
                        if (strid != 0)
                        {
                            var select = from s in npe.GongHuoFang where s.Id == strid select s;
                            foreach (var item in select)
                            {
                                npe.GongHuoFang.Remove(item);
                            }
                        }
                    }
                }
                npe.SaveChanges();
            }

        }
        #endregion

        public static string GetCellValue(WorkbookPart wbPart, Cell theCell)
        {
            string value = theCell.InnerText;
            //String value1 = theCell.CellValue.InnerText;
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




    }
}