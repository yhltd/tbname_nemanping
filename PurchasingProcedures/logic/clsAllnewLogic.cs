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
        #region 查询色号
        public List<Sehao> selectSehao()
        {
            using (nemanpingEntities3 can = new nemanpingEntities3())
            {
                List<Sehao> sehao = new List<Sehao>();
                var select = from s in can.Sehao select s;
                //DataTable dt = new DataTable();

                //dt.Columns.Add("Id", typeof(int));
                //dt.Columns.Add("Name", typeof(String));
                //dt.Columns.Add("SeHao1", typeof(String));
                foreach (var item in select)
                {
                    Sehao s = new Sehao();
                    s.Id = item.Id;
                    s.Name = item.Name;
                    s.SeHao1 = item.SeHao1;
                    sehao.Add(s);
                    //dt.Rows.Add(item.Id, item.Name, item.SeHao1);
                }
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

                    if (dr[0] is DBNull || Convert.ToInt32(dr[0])==0)
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
