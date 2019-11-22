﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using logic;
using System.IO;
using Spire.Xls;
namespace PurchasingProcedures
{
    public partial class ShengChengBiaoGeXuanZe : Form
    {
        protected string buttonType ;
        private string foldPath;
        private DataGridView dgv_ps;
        private DataGridView dgv_dh;
        private List<string> color;
        private int lie;
        protected GongNeng2 gn;
        private string STYLE;
        private string cdNo;
        private string foldPath2;
        private string imagePath;
        private string imagePath2;
        private string dayingType;
        public ShengChengBiaoGeXuanZe(string type)
        {
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            gn = new GongNeng2();
            buttonType = type;
            InitializeComponent();
        }
        public void ChangeExcel2Image(string filename)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(filename);
            Worksheet sheet = workbook.Worksheets[0];
            imagePath = Directory.GetCurrentDirectory() + "\\image.bmp";
            sheet.SaveToImage(imagePath); //图片后缀.bmp ,imagepath自己设置
        }
        public ShengChengBiaoGeXuanZe(string type, DataGridView dv1, DataGridView dv2, List<string> cl, int Lie, string style, string cdno)
        {
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            gn = new GongNeng2();
            dgv_ps = dv1;
            dgv_dh = dv2;
            color = cl;
            lie = Lie;
            STYLE = style;
            cdNo = cdno;
            buttonType = type;
            InitializeComponent();
        }
        private void CreateExcel(string path,string bctype )
        {
            DataTable dt = new DataTable();
            DataTable dt2 = new DataTable();
            if (bctype.Equals("配色"))
            {
                for (int i = 0; i < dgv_ps.Columns.Count; i++)
                {
                    if (!dgv_ps.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                    {
                        dt.Columns.Add(dgv_ps.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                    }
                }




                string C1str = "品名=货号=规格/幅宽";
                for (int i = 3; i < dgv_ps.ColumnCount; i++)
                {

                    C1str = C1str + "=" + dgv_ps.Columns[i].HeaderCell.Value.ToString();
                }
                dt.Rows.Add(C1str.Split('='));
                string C2str = "面料颜色= = ";
                for (int i = 0; i < color.Count; i++)
                {
                    if (!C2str.Contains(color[i]))
                    {
                        C2str = C2str + "=" + color[i];
                    }
                }
                dt.Rows.Add(C2str.Split('='));


                for (int i = 0; i < dgv_ps.Rows.Count; i++)
                {
                    string str = "";
                    //if (dgv_ps.Rows[i].Cells[6].Value != null)
                    //{
                    str = dgv_ps.Rows[i].Cells[0].Value + "=" + dgv_ps.Rows[i].Cells[1].Value + "=" + dgv_ps.Rows[i].Cells[2].Value;

                    for (int j = 3; j < lie+3; j++)
                    {

                        if (dgv_ps.Rows[i].Cells[j].Value != null)
                        {
                            str = str + "=" + dgv_ps.Rows[i].Cells[j].Value;
                        }

                    }
                    dt.Rows.Add(str.Split('='));
                    //}
                }

                foldPath = path + "\\配色表-" + STYLE + "-" + cdNo + ".xls";
            }
            else
            {


                for (int i = 0; i < dgv_dh.Columns.Count; i++)
                {
                    if (!dgv_dh.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                    {
                        dt2.Columns.Add(dgv_dh.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                    }
                }
                for (int i = 0; i < dgv_dh.Rows.Count; i++)
                {
                    if (dgv_dh.Rows[i].Cells[6].Value != null)
                    {
                        string str = "";
                        str = dgv_dh.Rows[i].Cells[0].Value + "=" + dgv_dh.Rows[i].Cells[1].Value + "=" + dgv_dh.Rows[i].Cells[2].Value + "=" + dgv_dh.Rows[i].Cells[3].Value + "=" + dgv_dh.Rows[i].Cells[4].Value + "=" + dgv_dh.Rows[i].Cells[5].Value + "=" + dgv_dh.Rows[i].Cells[6].Value;
                        dt2.Rows.Add(str.Split('='));
                    }
                }
                foldPath = path + "\\单耗-" + STYLE + "-" + cdNo + ".xls";
            }
            gn.SavePeiSeToExcel(dt, dt2, path, STYLE, cdNo);

        }
        public void DaYin(string Type) 
        {
            string path = Directory.GetCurrentDirectory();
            CreateExcel(path + "\\DaYingHuanCun", Type);
            if (foldPath != "")
            {

                ChangeExcel2Image(foldPath);  //利用Spire将excel转换成图片
                if (printDialog1.ShowDialog() == DialogResult.OK)
                {
                    printDocument1.Print();   //打印

                }
                File.Delete(foldPath);
                File.Delete(imagePath);
            }
            
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (buttonType.Equals("保存"))
            {
                try
                {
                    FolderBrowserDialog dialog = new FolderBrowserDialog();
                    dialog.Description = "请选择文件路径";
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        foldPath = dialog.SelectedPath;
                        CreateExcel(foldPath, "配色");
                        MessageBox.Show("保存成功！");
                    }
                }
                catch (Exception ex) { throw ex; }
            }
            else 
            {
                dayingType = "配色";
                DaYin("配色");
                
            }
        }

        private void ShengChengBiaoGeXuanZe_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (buttonType.Equals("保存"))
            {
                try
                {
                    FolderBrowserDialog dialog = new FolderBrowserDialog();
                    dialog.Description = "请选择文件路径";
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        foldPath = dialog.SelectedPath;
                        CreateExcel(foldPath, "单耗");
                        MessageBox.Show("保存成功！");
                    }
                }
                catch (Exception ex) { throw ex; }
            }
            else 
            {
                dayingType = "单耗";
                DaYin("单耗");
            }
            
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            #region 如果不需要截取图片，可以不用写以下代码
            GC.Collect();
            Graphics g = e.Graphics;
            //imagepath是指 excel转成的图片的路径
            using (Bitmap bitmap = new Bitmap(imagePath))
            {
                //如何截取自己摸索
                Rectangle newarea = new Rectangle();
                newarea.X = 0;
                newarea.Y = 50;
                newarea.Width = bitmap.Width;
                newarea.Height = bitmap.Height - 50;
                using (Bitmap newbitmap = bitmap.Clone(newarea, bitmap.PixelFormat))
                {
                    if (dayingType.Equals("配色"))
                    {
                        g.DrawImage(newbitmap, 0, 0, newbitmap.Width - 430, newbitmap.Height - 150);
                    }
                    else 
                    {
                        g.DrawImage(newbitmap, 0, 0, newbitmap.Width - 200, newbitmap.Height - 150);
                    }
                }
            }
            #endregion
        }
    }
}
