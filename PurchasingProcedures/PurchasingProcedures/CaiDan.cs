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
using clsBuiness;
using Spire.Xls;
using System.IO;

namespace PurchasingProcedures
{
    public partial class CaiDan : Form
    {
        //private PrintDocument pd = new PrintDocument();
        public string StyleId;
        public string chima;
        protected clsAllnewLogic cal;
        protected Definefactoryinput dfi;
        protected GongNeng2 gn2;
        protected List<clsBuiness.CaiDan> cd;
        protected List<JiaGongChang> jgc;
        protected string foldPath;
        protected string imagePath;
        public CaiDan(string style,string cmdp)
        {

            InitializeComponent();
            StyleId = style;
            chima = cmdp;
            cal = new clsAllnewLogic();
            dfi = new Definefactoryinput();
            gn2 = new GongNeng2();
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            cd = gn2.CreateCaiDan(style, cmdp);
            if (cd.Count ==0) 
            {
                MessageBox.Show("生成失败！ 原因:没有该Style的数据");
                this.Close();
            }
            //pd.PrinterSettings.PrinterName = "pdfFactory Pro";
            jgc = dfi.selectJiaGongChang().GroupBy(j => j.Name).Select(s => s.First()).ToList<JiaGongChang>();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void CaiDan_Load(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.CellValueChanged -= dataGridView1_CellValueChanged;
                txt_Style.Text = StyleId;
                txt_Label.Text = chima;
                this.txt_desc.Text = cd[0].DESC.ToString();
                this.txt_fabric.Text = cd[0].FABRIC.ToString();
                this.txt_jacket.Text = cd[0].Jacket.ToString();
                this.txt_pant.Text = cd[0].Pant.ToString();
                this.txt_shuoming.Text = cd[0].shuoming.ToString();
                cb_jgc.DataSource = jgc;
                cb_jgc.DisplayMember = "Name";
                cb_jgc.ValueMember = "id";
                txt_zhidan.Text = DateTime.Now.ToString() ;
                dataGridView1.CellValueChanged += dataGridView1_CellValueChanged;
            }
            catch (Exception ex) 
            {
                throw ex;
            }
            
        }

        private void dataGridView1_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            // 设定单元格的默认值
            e.Row.Cells["STYLE"].Value = StyleId;
        }

        private void 裁单输入ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try 
            {
                FolderBrowserDialog dialog = new FolderBrowserDialog();
                dialog.Description = "请选择文件路径";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    foldPath = dialog.SelectedPath;
                    CreateExcel(foldPath);
                    MessageBox.Show("生成成功！");
                }
            }
            catch (Exception ex) { throw ex; }
         
        }

        private void CreateExcel(string path)
        {
            DataTable dt = new DataTable();
            clsBuiness.CaiDan cd = new clsBuiness.CaiDan();
            cd.STYLE = txt_Style.Text;
            cd.LABEL = txt_Label.Text;
            cd.DESC = this.txt_desc.Text;
            cd.FABRIC = this.txt_fabric.Text;
            cd.Jacket = this.txt_jacket.Text;
            cd.Pant = this.txt_pant.Text;
            cd.shuoming = this.txt_shuoming.Text;
            cd.JiaGongchang = this.cb_jgc.Text;
            cd.MianLiao = this.txt_mianlioa.Text;
            cd.CaiDanHao = this.txt_CaidanNo.Text;
            cd.ZhiDanRiqi = this.txt_zhidan.Text;
            cd.JiaoHuoRiqi = this.txt_jiaohuo.Text;
            cd.RN_NO = this.txt_RN.Text;
            dt.Columns.Add("id", typeof(int));
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                {
                    dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                }
            }
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Cells[6].Value != null)
                {
                    dt.Rows.Add(dataGridView1.Rows[i].Cells[0].Value, dataGridView1.Rows[i].Cells[1].Value, dataGridView1.Rows[i].Cells[2].Value, dataGridView1.Rows[i].Cells[3].Value, dataGridView1.Rows[i].Cells[4].Value, dataGridView1.Rows[i].Cells[5].Value, dataGridView1.Rows[i].Cells[6].Value, dataGridView1.Rows[i].Cells[7].Value, dataGridView1.Rows[i].Cells[8].Value, dataGridView1.Rows[i].Cells[9].Value, dataGridView1.Rows[i].Cells[10].Value, dataGridView1.Rows[i].Cells[11].Value, dataGridView1.Rows[i].Cells[12].Value, dataGridView1.Rows[i].Cells[13].Value, dataGridView1.Rows[i].Cells[14].Value, dataGridView1.Rows[i].Cells[15].Value, dataGridView1.Rows[i].Cells[16].Value, dataGridView1.Rows[i].Cells[17].Value, dataGridView1.Rows[i].Cells[18].Value, dataGridView1.Rows[i].Cells[19].Value, dataGridView1.Rows[i].Cells[20].Value, dataGridView1.Rows[i].Cells[21].Value, dataGridView1.Rows[i].Cells[22].Value, dataGridView1.Rows[i].Cells[23].Value, dataGridView1.Rows[i].Cells[24].Value, dataGridView1.Rows[i].Cells[25].Value, dataGridView1.Rows[i].Cells[26].Value, dataGridView1.Rows[i].Cells[27].Value, dataGridView1.Rows[i].Cells[28].Value, dataGridView1.Rows[i].Cells[29].Value, dataGridView1.Rows[i].Cells[30].Value, dataGridView1.Rows[i].Cells[31].Value, dataGridView1.Rows[i].Cells[32].Value, dataGridView1.Rows[i].Cells[33].Value, dataGridView1.Rows[i].Cells[34].Value, dataGridView1.Rows[i].Cells[35].Value, dataGridView1.Rows[i].Cells[36].Value, dataGridView1.Rows[i].Cells[37].Value, dataGridView1.Rows[i].Cells[38].Value, dataGridView1.Rows[i].Cells[39].Value, dataGridView1.Rows[i].Cells[40].Value, dataGridView1.Rows[i].Cells[41].Value, dataGridView1.Rows[i].Cells[42].Value, dataGridView1.Rows[i].Cells[43].Value);
                }
            }
            gn2.CDEXCEL(dt, cd, path);
            foldPath = path + "\\裁单表-" + cd.CaiDanHao + ".xls";
        }
        public void ChangeExcel2Image(string filename)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(filename);
            Worksheet sheet = workbook.Worksheets[0];
            imagePath = Directory.GetCurrentDirectory() + "\\image.bmp";

            sheet.SaveToImage(imagePath); //图片后缀.bmp ,imagepath自己设置
        }
        private void 面辅料订购ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //try 
            //{
                //isprint = false;
                //裁单输入ToolStripMenuItem_Click(sender, e); //使用NPOI生成excel
                string path = Directory.GetCurrentDirectory();
                CreateExcel(path + "\\DaYingHuanCun");
                if (foldPath != "")
                {
                    //isprint = false;
                    ChangeExcel2Image(foldPath);  //利用Spire将excel转换成图片
                    if (printDialog1.ShowDialog() == DialogResult.OK)
                    {
                        printDocument1.Print();   //打印

                    }
                    File.Delete(foldPath);
                    File.Delete(imagePath);
                } 
            //}
            //catch (Exception ex) { throw ex; }
            

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
                newarea.Height = bitmap.Height -50;
                using (Bitmap newbitmap = bitmap.Clone(newarea, bitmap.PixelFormat))
                {
                    g.DrawImage(newbitmap, 0, 0, newbitmap.Width-430, newbitmap.Height - 150);
                }
            }
            #endregion
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            //if (e.RowIndex >= 0)
            //{
            //    int sum = 0;
            //    for (int i = 7; i < dataGridView1.Columns.Count; i++)
            //    {
               
            //            sum = sum + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[i].Value);
                
            //    }
            //    dataGridView1.Rows[e.RowIndex].Cells["Column43"].Value = sum;
            //}
        }

        private void dataGridView1_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {

            dataGridView1.Rows[e.RowIndex].Cells["Column43"].Value = Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column7"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column8"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column9"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column10"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column11"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column12"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column13"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column14"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column15"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column16"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column17"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column18"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column19"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column20"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column21"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column22"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column23"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column24"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column25"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column26"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column27"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column28"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column29"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column30"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column31"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column32"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column33"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column34"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column35"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column36"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column37"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column38"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column39"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column40"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column41"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column42"].Value);
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
    }
}
