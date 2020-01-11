using System;
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
        protected string buttonType;
        private string foldPath;
        private DataGridView dgv_ps;
        private DataGridView dataGridView1;
        private DataGridView dgv_dh;
        private DataGridView dataGridView2;
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
        public ShengChengBiaoGeXuanZe(string type, DataGridView dv1, DataGridView dv2, DataGridView dv3, List<string> cl, int Lie, string style, string cdno, DataGridView dataGridView21)
        {
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            gn = new GongNeng2();
            dgv_ps = dv1;
            dataGridView1 = dv3;
            dgv_dh = dv2;
            dataGridView2 = dataGridView21;

            color = cl;
            lie = Lie;
            STYLE = style;
            cdNo = cdno;
            buttonType = type;
            InitializeComponent();
        }
        private void CreateExcel(string path, string bctype)
        {
            DataTable dt = new DataTable();
            DataTable dt2 = new DataTable();
            DataTable dt3 = new DataTable();
            DataTable dt4 = new DataTable();
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

                    for (int j = 3; j < lie + 3; j++)
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
            else if (bctype.Equals("单耗"))
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
            else if (bctype.Equals("核定成本"))
            {

                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                    {
                        dt3.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                    }
                }
                dt3.Columns.Add("合计", typeof(string));
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1.Rows[i].Cells[0].Value != null)
                    {
                        string str = "";
                        if (!dataGridView1.Rows[i].Cells[1].Value.ToString().Equals(string.Empty) && !dataGridView1.Rows[i].Cells[2].Value.Equals(string.Empty))
                        {
                            str = dataGridView1.Rows[i].Cells[0].Value + "=" + dataGridView1.Rows[i].Cells[1].Value + "=" + dataGridView1.Rows[i].Cells[2].Value + "=" + (Convert.ToInt32(dataGridView1.Rows[i].Cells[1].Value) * Convert.ToInt32(dataGridView1.Rows[i].Cells[2].Value));
                        }
                        else
                        {
                            str = dataGridView1.Rows[i].Cells[0].Value + "=" + dataGridView1.Rows[i].Cells[1].Value + "=" + dataGridView1.Rows[i].Cells[2].Value + "=" + 0;
                        }
                        dt3.Rows.Add(str.Split('='));
                    }
                }
                foldPath = path + "\\核定成本-" + STYLE + "-" + cdNo + ".xls";

            }
            else if (bctype.Equals("核算表"))
            {
           
                for (int i = 0; i < dataGridView2.Columns.Count; i++)
                {
                    if (!dataGridView2.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                    {
                        dt4.Columns.Add(dataGridView2.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                    }
                }
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    if (dataGridView2.Rows[i].Cells[0].Value != null)
                    {
                        string str = "";
                        //if (!dataGridView2.Rows[i].Cells[1].Value.ToString().Equals(string.Empty) && !dataGridView2.Rows[i].Cells[2].Value.Equals(string.Empty))
                        //{
                        //    str = dataGridView2.Rows[i].Cells[0].Value + "=" + dataGridView2.Rows[i].Cells[1].Value + "=" + dataGridView2.Rows[i].Cells[2].Value + "=" + (Convert.ToInt32(dataGridView1.Rows[i].Cells[1].Value) * Convert.ToInt32(dataGridView2.Rows[i].Cells[2].Value));
                        //}
                        //else
                        //{
                        //    str = dataGridView2.Rows[i].Cells[0].Value + "=" + dataGridView2.Rows[i].Cells[1].Value + "=" + dataGridView2.Rows[i].Cells[2].Value + "=" + 0;
                        //}

                        for (int j = 0; j < dataGridView2.Columns.Count; j++)
                        {
                            if (dataGridView2.Rows[i].Cells[j].Value != null)
                            {
                                if (str != "")
                                    str = str + "=" + dataGridView2.Rows[i].Cells[j].Value;//品名=货号=规格/幅宽=61605C1=61607C1=61609C1=61601C1=61627C1=61634C1
                                else
                                    str = dataGridView2.Rows[i].Cells[j].Value.ToString();//品名=货号=规格/幅宽=61605C1=61607C1=61609C1=61601C1=61627C1=61634C1
                            }
                        }
                        dt4.Rows.Add(str.Split('='));
                    }
                }
            }
            gn.SavePeiSeToExcel(dt, dt2, dt3, path, STYLE, cdNo, dt4);

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
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                MessageBox.Show("建议横向打印");
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
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
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
                    else if (dayingType.Equals("单耗"))
                    {
                        g.DrawImage(newbitmap, 0, 0, newbitmap.Width - 200, newbitmap.Height - 150);
                    }
                    else
                    {
                        g.DrawImage(newbitmap, 0, 0, newbitmap.Width + 120, newbitmap.Height - 100);
                    }
                }
            }
            #endregion
        }

        private void button3_Click(object sender, EventArgs e)
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
                        CreateExcel(foldPath, "核定成本");
                        MessageBox.Show("保存成功！");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                dayingType = "核定成本";
                DaYin("核定成本");
            }
        }

        private void button4_Click(object sender, EventArgs e)
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
                        CreateExcel(foldPath, "核算表");
                        MessageBox.Show("保存成功！");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                dayingType = "核算表";
                DaYin("核算表");
            }
        }
    }
}
