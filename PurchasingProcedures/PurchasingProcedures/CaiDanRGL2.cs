using clsBuiness;
using logic;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PurchasingProcedures
{
    public partial class CaiDanRGL2 : Form
    {
        protected GongNeng2 gn2;

        protected string foldPath;
        private string sty;
        private string cmd;
        protected Definefactoryinput dfi;
        protected List<JiaGongChang> jgc;

        protected List<clsBuiness.CaiDan> cd;
        private Form FM;
        private Form frm;
        public CaiDanRGL2(string style, string cmdp, Form f)
        {
            gn2 = new GongNeng2();
            sty = style;
            cmd = cmdp;
            frm = f;
            FM = f;
            cd = gn2.CreateCaiDan(style, cmdp);
            dfi = new Definefactoryinput();
            if (cd.Count == 0)
            {
                MessageBox.Show("生成失败！ 原因:没有该Style的数据");
                this.Close();
            }
            //pd.PrinterSettings.PrinterName = "pdfFactory Pro";
            jgc = dfi.selectJiaGongChang().GroupBy(j => j.Name).Select(s => s.First()).ToList<JiaGongChang>();

            InitializeComponent();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (txt_CaidanNo.Text != null && !txt_CaidanNo.Text.Equals(string.Empty))
                {
                    FolderBrowserDialog dialog = new FolderBrowserDialog();
                    dialog.Description = "请选择文件路径";
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        foldPath = dialog.SelectedPath;
                        CreateExcel(foldPath);
                        MessageBox.Show("生成成功！");
                    }
                    Form1_Load(sender, e);


                }
                else
                {
                    MessageBox.Show("裁单号不能为空!");
                }
            }


            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
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
            //dt.Columns.Add("id", typeof(int));
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                //if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                //{
                dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                //}
            }
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Cells[6].Value != null)
                {
                    dt.Rows.Add(dataGridView1.Rows[i].Cells[0].Value, dataGridView1.Rows[i].Cells[1].Value, dataGridView1.Rows[i].Cells[2].Value, dataGridView1.Rows[i].Cells[3].Value, dataGridView1.Rows[i].Cells[4].Value, dataGridView1.Rows[i].Cells[5].Value, dataGridView1.Rows[i].Cells[6].Value, dataGridView1.Rows[i].Cells[7].Value, dataGridView1.Rows[i].Cells[8].Value, dataGridView1.Rows[i].Cells[9].Value, dataGridView1.Rows[i].Cells[10].Value, dataGridView1.Rows[i].Cells[11].Value, dataGridView1.Rows[i].Cells[12].Value, dataGridView1.Rows[i].Cells[13].Value, dataGridView1.Rows[i].Cells[14].Value, dataGridView1.Rows[i].Cells[15].Value, dataGridView1.Rows[i].Cells[16].Value, dataGridView1.Rows[i].Cells[17].Value, dataGridView1.Rows[i].Cells[18].Value, dataGridView1.Rows[i].Cells[19].Value, dataGridView1.Rows[i].Cells[20].Value, dataGridView1.Rows[i].Cells[21].Value, dataGridView1.Rows[i].Cells[22].Value, dataGridView1.Rows[i].Cells[23].Value, dataGridView1.Rows[i].Cells[24].Value, dataGridView1.Rows[i].Cells[25].Value, dataGridView1.Rows[i].Cells[26].Value, dataGridView1.Rows[i].Cells[27].Value, dataGridView1.Rows[i].Cells[28].Value, dataGridView1.Rows[i].Cells[29].Value, dataGridView1.Rows[i].Cells[30].Value, dataGridView1.Rows[i].Cells[31].Value, dataGridView1.Rows[i].Cells[32].Value, dataGridView1.Rows[i].Cells[33].Value, dataGridView1.Rows[i].Cells[34].Value, dataGridView1.Rows[i].Cells[35].Value, dataGridView1.Rows[i].Cells[36].Value, dataGridView1.Rows[i].Cells[37].Value, dataGridView1.Rows[i].Cells[38].Value, dataGridView1.Rows[i].Cells[39].Value, dataGridView1.Rows[i].Cells[40].Value, dataGridView1.Rows[i].Cells[41].Value, dataGridView1.Rows[i].Cells[42].Value, dataGridView1.Rows[i].Cells[43].Value);
                }
            }
            gn2.rgl2EXCEL(dt, cd, path);
            foldPath = path + "\\裁单表-" + cd.CaiDanHao + ".xls";
        }


        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.None;//设置dava宽度

    

            #region RGL2
            DataTable dt1 = new DataTable();
            dt1.Columns.Add("Id", typeof(int));
            dt1.Columns.Add("面料", typeof(string));
            dt1.Columns.Add("款式", typeof(string));
            dt1.Columns["款式"].ReadOnly = true;
            dt1.Columns.Add("货号", typeof(string));
            dt1.Columns.Add("颜色", typeof(string));
            dt1.Columns.Add("颜色编号", typeof(string));
            dt1.Columns.Add("裤子", typeof(string));
            dt1.Columns.Add("28", typeof(string));
            dt1.Columns.Add("30", typeof(string));
            dt1.Columns.Add("32", typeof(string));
            dt1.Columns.Add("34", typeof(string));
            dt1.Columns.Add("36", typeof(string));
            dt1.Columns.Add("38", typeof(string));
            dt1.Columns.Add("40", typeof(string));
            dt1.Columns.Add("42", typeof(string));
            dt1.Columns.Add("44", typeof(string));
            dt1.Columns.Add("46", typeof(string));
            dt1.Columns.Add("48", typeof(string));
            dt1.Columns.Add("50", typeof(string));
            dt1.Columns.Add("52", typeof(string));
            dt1.Columns.Add("54", typeof(string));
            dt1.Columns.Add("56", typeof(string));
            dt1.Columns.Add(" 30", typeof(string));
            dt1.Columns.Add(" 32", typeof(string));
            dt1.Columns.Add(" 34", typeof(string));
            dt1.Columns.Add(" 36", typeof(string));
            dt1.Columns.Add(" 38", typeof(string));
            dt1.Columns.Add(" 40", typeof(string));
            dt1.Columns.Add(" 42", typeof(string));
            dt1.Columns.Add(" 44", typeof(string));
            dt1.Columns.Add(" 46", typeof(string));
            dt1.Columns.Add(" 48", typeof(string));
            dt1.Columns.Add(" 50", typeof(string));
            dt1.Columns.Add(" 52", typeof(string));
            dt1.Columns.Add(" 54", typeof(string));
            dt1.Columns.Add(" 56", typeof(string));
            dt1.Columns.Add(" 58", typeof(string));
            dt1.Columns.Add(" 60", typeof(string));
            dt1.Columns.Add(" 62", typeof(string));
            dt1.Columns.Add("34 ", typeof(string));
            dt1.Columns.Add("36 ", typeof(string));
            dt1.Columns.Add("38 ", typeof(string));
            dt1.Columns.Add("40 ", typeof(string));
            dt1.Columns.Add("Sub Total: ", typeof(string));
            dataGridView1.DataSource = dt1;
            DataGridViewHelper rowMergeView = new DataGridViewHelper(dataGridView1);
            dataGridView1.Columns["Id"].Visible = false;
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(1, 1, "LOT#"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(2, 1, "STYLE"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(3, 1, "ART"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(4, 1, "COLOR"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(5, 1, "COLOR#"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(6, 1, "上衣"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(7, 1, "34R"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(8, 1, "36R"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(9, 1, "38R"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(10, 1, "40R"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(11, 1, "42R"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(12, 1, "44R"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(13, 1, "46R"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(14, 1, "48R"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(15, 1, "50R"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(16, 1, "52R"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(17, 1, "54R"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(18, 1, "56R"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(19, 1, "58R"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(20, 1, "60R"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(21, 1, "62R"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(22, 1, "36L"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(23, 1, "38L"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(24, 1, "40L"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(25, 1, "42L"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(26, 1, "44L"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(27, 1, "46L"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(28, 1, "48L"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(29, 1, "50L"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(30, 1, "52L"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(31, 1, "54L"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(32, 1, "56L"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(33, 1, "58L"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(34, 1, "60L"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(35, 1, "62L"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(36, 1, "64L"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(37, 1, "36S"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(38, 1, "38S"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(39, 1, "40S"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(40, 1, "42S"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(41, 1, "44S"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(42, 1, "46S"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(43, 1, "订单合计"));

            #endregion


       




            dataGridView1.CellValueChanged -= dataGridView1_CellValueChanged;

            txt_Style.Text = sty;
            txt_Label.Text = cmd;
            txt_CaidanNo.SelectedIndexChanged -= txt_CaidanNo_SelectedIndexChanged;
            this.txt_desc.Text = cd[0].DESC.ToString();
            this.txt_fabric.Text = cd[0].FABRIC.ToString();
            this.txt_jacket.Text = cd[0].Jacket.ToString();
            this.txt_pant.Text = cd[0].Pant.ToString();
            this.txt_shuoming.Text = cd[0].shuoming.ToString();
            dataGridView1.ColumnHeadersHeight = 35;
            txt_CaidanNo.SelectedIndexChanged += txt_CaidanNo_SelectedIndexChanged;
            cb_jgc.DataSource = jgc;
            cb_jgc.DisplayMember = "Name";
            cb_jgc.ValueMember = "id";
            txt_zhidan.Text = DateTime.Now.ToString();
            List<clsBuiness.CaiDan_RGL2> caidan = gn2.selectCaiDanRGL2("").GroupBy(g => g.CaiDanHao).Select(s => s.First()).ToList<clsBuiness.CaiDan_RGL2>();
            clsBuiness.CaiDan_RGL2 c = new clsBuiness.CaiDan_RGL2()
            {
                CaiDanHao = " "
            };
            caidan.Add(c);
            txt_CaidanNo.DataSource = caidan;
            txt_CaidanNo.DisplayMember = "CaiDanHao";
            txt_CaidanNo.ValueMember = "id";
            if (txt_CaidanNo.FindString(" ") >= 0)
            {
                txt_CaidanNo.SelectedIndex = txt_CaidanNo.FindString(" ");
            }
            //dataGridView1.CellValueChanged += dataGridView1_CellValueChanged;


            //dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            //   dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            //dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            // dataGridView1.Columns[7].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            resizedava_cloumn(dataGridView1);//设置Dave 宽度

            //for (int j = 0; j < 43; j++)
            //{
            //    //if (j < dataGridView1.ColumnCount - Convert.ToInt32(toolStripComboBox4.Text))
            //    dataGridView1.Columns[j].Width = 500;
            //    }

        }

        private void resizedava_cloumn(DataGridView dataGridView)
        {
            dataGridView.Columns[1].Width = 60;
            dataGridView.Columns[2].Width = 60;
            dataGridView.Columns[3].Width = 60;
            dataGridView.Columns[4].Width = 70;
            dataGridView.Columns[5].Width = 60;
            dataGridView.Columns[6].Width = 30;
            dataGridView.Columns[7].Width = 25;
            dataGridView.Columns[8].Width = 25;
            dataGridView.Columns[9].Width = 25;
            dataGridView.Columns[10].Width = 25;
            dataGridView.Columns[11].Width = 25;
            dataGridView.Columns[12].Width = 25;
            dataGridView.Columns[13].Width = 25;
            dataGridView.Columns[14].Width = 25;
            dataGridView.Columns[15].Width = 25;
            dataGridView.Columns[16].Width = 25;
            dataGridView.Columns[17].Width = 25;
            dataGridView.Columns[18].Width = 25;
            dataGridView.Columns[19].Width = 25;
            dataGridView.Columns[20].Width = 25;
            dataGridView.Columns[21].Width = 25;
            dataGridView.Columns[22].Width = 25;
            dataGridView.Columns[23].Width = 25;
            dataGridView.Columns[24].Width = 25;
            dataGridView.Columns[25].Width = 25;
            dataGridView.Columns[26].Width = 25;
            dataGridView.Columns[27].Width = 25;
            dataGridView.Columns[28].Width = 25;
            dataGridView.Columns[29].Width = 25;
            dataGridView.Columns[30].Width = 25;
            dataGridView.Columns[31].Width = 25;
            dataGridView.Columns[32].Width = 25;
            dataGridView.Columns[33].Width = 25;
            dataGridView.Columns[34].Width = 25;
            dataGridView.Columns[35].Width = 25;
            dataGridView.Columns[36].Width = 25;
            dataGridView.Columns[37].Width = 25;
            dataGridView.Columns[38].Width = 25;
            dataGridView.Columns[39].Width = 25;
            dataGridView.Columns[40].Width = 25;
            dataGridView.Columns[41].Width = 25;
            dataGridView.Columns[42].Width = 25;
            dataGridView.Columns[43].Width = 60;




        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            try
            {
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
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);


            }
        }
        public void ChangeExcel2Image(string filename)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(filename);
            Worksheet sheet = workbook.Worksheets[0];
            imagePath = Directory.GetCurrentDirectory() + "\\image.bmp";

            sheet.SaveToImage(imagePath); //图片后缀.bmp ,imagepath自己设置
        }

        public string imagePath { get; set; }

        private void button2_Click(object sender, EventArgs e)
        {
            {
                try
                {
                    MianFuLiaoDingGou mfldg = new MianFuLiaoDingGou(txt_mianlioa.Text, txt_Style.Text, cb_jgc.Text, txt_CaidanNo.Text, FM);
                    if (!mfldg.IsDisposed)
                    {
                        if (!HaveOpened(FM, mfldg.Name))
                        {
                            mfldg.MdiParent = FM;
                            mfldg.Show();
                        }
                        else
                        {
                            mfldg.TopMost = true;
                        }
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);

                }
            }
        }
        private bool HaveOpened(Form _monthForm, string _childrenFormName)
        {
            //查看窗口是否已经被打开
            bool bReturn = false;
            for (int i = 0; i < _monthForm.MdiChildren.Length; i++)
            {
                if (_monthForm.MdiChildren[i].Name == _childrenFormName)
                {
                    _monthForm.MdiChildren[i].BringToFront();//将控件带到 Z 顺序的前面。
                    bReturn = true;
                    break;
                }
            }
            return bReturn;
        }


        public Form F { get; set; }
        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if (e.ColumnIndex >= 7 && e.ColumnIndex <= 42 /*绑定数据源中列的序号*/)
                {
                    //dataGridView1.CellFormatting -= dataGridView1_CellFormatting;
                    double sum = 0;
                    for (int i = 7; i <= 42; i++)
                    {
                        if (dataGridView1.Rows[e.RowIndex].Cells[i].Value != null && !dataGridView1.Rows[e.RowIndex].Cells[i].Value.ToString().Equals(string.Empty) && IsNumberic(dataGridView1.Rows[e.RowIndex].Cells[i].Value.ToString()))
                        {
                            sum = sum + Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[i].Value);
                        }
                        //dataGridView1.Rows[e.RowIndex].Cells["Column43"].Value = Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column7"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column8"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column9"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column10"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column11"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column12"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column13"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column14"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column15"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column16"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column17"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column18"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column19"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column20"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column21"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column22"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column23"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column24"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column25"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column26"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column27"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column28"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column29"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column30"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column31"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column32"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column33"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column34"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column35"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column36"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column37"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column38"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column39"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column40"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column41"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column42"].Value);
                    }

                    dataGridView1.Rows[e.RowIndex].Cells[43].Value = sum;
                    //dataGridView1.CellFormatting += dataGridView1_CellFormatting;  
                    //dataGridView1.Rows[0].Cells["Sub_Total"].Value = sum.ToString();
                }
            }


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
        private void txt_CaidanNo_SelectedIndexChanged(object sender, EventArgs e)
        {
            {
                try
                {
                    DataTable dt = new DataTable();
                    List<clsBuiness.CaiDan_RGL2> cdlist = new List<CaiDan_RGL2>();
                    if (!txt_CaidanNo.Text.Equals(string.Empty))
                    {
                        cdlist = gn2.selectCaiDanRGL2(txt_CaidanNo.Text);
                    }

                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        //if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                        //{
                        dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                        //}
                    }
                    foreach (clsBuiness.CaiDan_RGL2 s in cdlist)
                    {
                        dt.Rows.Add(s.Id, s.LOT, s.STYLE, s.ART, s.COLOR, s.COLORID, s.JACKET_PANT, s.C34R, s.C36R, s.C38R, s.C40R, s.C42R, s.C44R, s.C46R, s.C48R, s.C50R, s.C52R, s.C54R, s.C56R, s.C58R, s.C60R, s.C62R, s.C36L, s.C38L, s.C40L, s.C42L, s.C44L, s.C46L, s.C48L, s.C50L, s.C52L, s.C54L, s.C56L, s.C58L, s.C60L, s.C62L, s.C34S, s.C36S, s.C38S, s.C40S, s.C42S, s.C44S, s.C46S, s.Sub_Total);
                    }

                    //foreach (clsBuiness.CaiDan s in cdlist)
                    //{
                    ////for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    ////{
                    //    //if (dataGridView1.Rows[i].Cells[6].Value != null)
                    //    //{
                    //                    dt.Rows.Add(s.Id, s.LOT, s.STYLE, s.ART, s.COLOR, s.COLORID, s.JACKET_PANT, s.C34R, s.C36R, s.C38R, s.C40R,s.C42R, s.C44R, s.C46R, s.C48R, s.C50R, s.C52R, s.C54R, s.C56R, s.C58R, s.C60R, s.C62R, s.C36L, s.C38L, s.C40L, s.C42L, s.C44L, s.C46L, s.C48L, s.C50L, s.C52L, s.C54L, s.C56L, s.C58L, s.C60L, s.C62L, s.C34S, s.C36S, s.C38S, s.C40S, s.C42S, s.C44S, s.C46S, s.Sub_Total);

                    //    //}
                    //}
                    if (cdlist != null && cdlist.Count > 0)
                    {
                        txt_Style.Text = cdlist[0].STYLE;
                        txt_Label.Text = cdlist[0].LABEL;
                        this.txt_desc.Text = cdlist[0].DESC.ToString();
                        this.txt_fabric.Text = cdlist[0].FABRIC.ToString();
                        this.txt_jacket.Text = cdlist[0].Jacket.ToString();
                        this.txt_pant.Text = cdlist[0].Pant.ToString();
                        this.txt_shuoming.Text = cdlist[0].shuoming.ToString();
                        this.txt_jiaohuo.Text = cdlist[0].JiaoHuoRiqi.ToString();
                        txt_zhidan.Text = DateTime.Now.ToLongDateString().ToString();
                        txt_RN.Text = cdlist[0].RN_NO.ToString();
                        txt_mianlioa.Text = cdlist[0].MianLiao.ToString();
                    }
                    dataGridView1.DataSource = dt;
                    dataGridView1.CellValueChanged += dataGridView1_CellValueChanged;

                }
                catch (Exception EX) { MessageBox.Show(EX.Message); }


            }
        }
        private bool IsNumberic(string oText)
        {
            try
            {
                int var1 = Convert.ToInt32(oText);
                return true;
            }
            catch
            {
                return false;
            }
        }
        private void dataGridView1_CellValueChanged_1(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if (e.ColumnIndex >= 7 && e.ColumnIndex <= 42 /*绑定数据源中列的序号*/)
                {
                    //dataGridView1.CellFormatting -= dataGridView1_CellFormatting;
                    double sum = 0;
                    for (int i = 7; i <= 42; i++)
                    {
                        if (dataGridView1.Rows[e.RowIndex].Cells[i].Value != null && !dataGridView1.Rows[e.RowIndex].Cells[i].Value.ToString().Equals(string.Empty) && IsNumberic(dataGridView1.Rows[e.RowIndex].Cells[i].Value.ToString()))
                        {
                            sum = sum + Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[i].Value);
                        }
                        //dataGridView1.Rows[e.RowIndex].Cells["Column43"].Value = Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column7"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column8"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column9"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column10"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column11"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column12"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column13"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column14"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column15"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column16"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column17"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column18"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column19"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column20"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column21"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column22"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column23"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column24"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column25"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column26"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column27"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column28"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column29"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column30"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column31"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column32"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column33"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column34"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column35"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column36"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column37"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column38"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column39"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column40"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column41"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column42"].Value);
                    }

                    dataGridView1.Rows[e.RowIndex].Cells[43].Value = sum;
                    //dataGridView1.CellFormatting += dataGridView1_CellFormatting;
                }
            }
        }

        private void dataGridView1_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            // 设定单元格的默认值
            e.Row.Cells["款式"].Value = sty;
        }
    }
}

