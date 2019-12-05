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
    public class Definefactoryinput
    {
        //public bool Login(string name, string pwd)
        //{
        //    using (nemanpingEntities3 can = new nemanpingEntities3())
        //    {
        //        var user = can.UserTable.First(u => u.Name.Equals(name) && u.Pwd.Equals(pwd));
        //        return true;
        //    }

        //}
        #region 查询加工厂信息
        public List<JiaGongChang> selectJiaGongChang()
        {
            using (nemanpingEntities3 can = new nemanpingEntities3())
            {
                List<JiaGongChang> JiaGongChang = new List<JiaGongChang>();
                var select = from s in can.JiaGongChang select s;
                foreach (var item in select)
                {
                    JiaGongChang s = new JiaGongChang();
                    s.id = item.id;
                    s.Name = item.Name;
                    s.Address = item.Address;
                    s.Lianxiren = item.Lianxiren;
                    s.Phone = item.Phone;
                    s.ZengZhiShui = item.ZengZhiShui;
                    s.Kaihuhang = item.Kaihuhang;
                    s.Zhanghao = item.Zhanghao;
                    JiaGongChang.Add(s);
                }
                return JiaGongChang;
            }
        }

        #endregion
        #region 添加工厂信息
        public void insertJiaGongChang(DataTable dt)
        {
            using (nemanpingEntities3 can = new nemanpingEntities3())
            {
                foreach (DataRow dr in dt.Rows)
                {

                    if (dr[0] is DBNull || Convert.ToInt32(dr[0]) == 0)
                    {
                        JiaGongChang insets = new JiaGongChang()
                        {
                            Name = dr[1].ToString(),
                            Address = dr[2].ToString(),
                            Lianxiren = dr[3].ToString(),
                            Phone = dr[4].ToString(),
                            ZengZhiShui = dr[5].ToString(),
                            Kaihuhang = dr[6].ToString(),
                            Zhanghao = dr[7].ToString()

                        };
                        can.JiaGongChang.Add(insets);

                    }
                    else
                    {
                        int id = Convert.ToInt32(dr[0]);
                        var select = from sc in can.JiaGongChang where sc.id == id select sc;
                        var target = select.FirstOrDefault<JiaGongChang>();
                        target.Name = dr[1].ToString();
                        target.Address = dr[2].ToString();
                        target.Lianxiren = dr[3].ToString();
                        target.Phone = dr[4].ToString();
                        target.ZengZhiShui = dr[5].ToString();
                        target.Kaihuhang = dr[6].ToString();
                        target.Zhanghao = dr[7].ToString();

                    }
                }
                can.SaveChanges();

            }
        }
        #endregion
        #region 读取工厂
        public List<JiaGongChang> readerJiaGongChangExcel(string fileName)
        {
            List<JiaGongChang> list = new List<JiaGongChang>();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart wbPart = document.WorkbookPart;
                List<Sheet> sheets = wbPart.Workbook.Descendants<Sheet>().ToList();
                var versionSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                WorksheetPart worksheetPart = (WorksheetPart)wbPart.GetPartById(versionSheet.Id);
                int rowindex = 0;
                int insertpd = 0;
                JiaGongChang s = new JiaGongChang();
                foreach (Row row in worksheetPart.Worksheet.Descendants<Row>())
                {
                    if (rowindex < 1)
                    {
                        rowindex++;
                        continue;
                    }
                    
                    foreach (Cell cell in row)
                    {
                        string rev = cell.CellReference.Value;
                        //if (rev.StartsWith("B"))
                        //{
                            if (rev.StartsWith("A")) 
                            {
                                s.Name = GetCellValue(wbPart, cell);
                            }
                            if (rev.StartsWith("B"))
                            {
                                s.Address = GetCellValue(wbPart, cell);
                            }
                            if (rev.StartsWith("C"))
                            {
                                s.Lianxiren = GetCellValue(wbPart, cell);
                            }
                            if (rev.StartsWith("D"))
                            {
                                s.Phone = GetCellValue(wbPart, cell);
                            }
                            if (rev.StartsWith("G"))
                            {
                                s.ZengZhiShui = GetCellValue(wbPart, cell);
                            }
                            if (rev.StartsWith("E"))
                            {
                                s.Kaihuhang = GetCellValue(wbPart, cell);
                            }
                            if (rev.StartsWith("F"))
                            {
                                s.Zhanghao = GetCellValue(wbPart, cell);
                                insertpd = 1;
                            }
                        //}
                        

                    }
                    if (insertpd == 1) 
                    {
                        
                        list.Add(s);
                        insertpd = 0;
                        s = new JiaGongChang();
                    }
                }
                return list;
            }
        }
        #endregion
        #region 删除加工厂信息
        public void deleteJaGongChang(List<int> id)
        {
            using (nemanpingEntities3 npe = new nemanpingEntities3())
            {
                foreach (int strid in id)
                {
                    if (strid != null)
                    {
                        if (strid != 0)
                        {
                            var select = from s in npe.JiaGongChang where s.id == strid select s;
                            foreach (var item in select)
                            {
                                npe.JiaGongChang.Remove(item);
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

    }
}
