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
    public partial class frmCaiDanD_PANT : Form
    {
        protected GongNeng2 gn2;
        protected string foldPath;
        protected string imagePath;
        public string StyleId;
        public string chima;
        //private Form FM;
        protected Definefactoryinput dfi;
        protected List<JiaGongChang> jgc;
        protected List<clsBuiness.CaiDan> cd;
        //private string kuanshi;
        //private string cmdp;
        private Form fm;
        public clsAllnewLogic cal = new clsAllnewLogic();
        public frmCaiDanD_PANT(string style, string cmdp, Form f)
        {
            //kuanshi = ks;
            //cmdp = cm;
            fm = f;
            gn2 = new GongNeng2();
            dfi = new Definefactoryinput();
            StyleId = style;
            chima = cmdp;
            cd = gn2.CreateCaiDan(style, cmdp);
            jgc = dfi.selectJiaGongChang().GroupBy(j => j.Name).Select(s => s.First()).ToList<JiaGongChang>();
            InitializeComponent();
        }
        public frmCaiDanD_PANT()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (txt_CaidanNo.Text != null && !txt_CaidanNo.Text.Equals(string.Empty))
                {
                    FolderBrowserDialog dialog = new FolderBrowserDialog();
                    //dialog.Description = "请选择文件路径";
                    //if (dialog.ShowDialog() == DialogResult.OK)
                    //{
                        foldPath =Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Result");
                        CreateExcel(foldPath);
                        MessageBox.Show("生成成功！");
                    //}
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
            string foldPath1 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Result");
            //string ZFCEPath = Path.Combine(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System"), "");
            System.Diagnostics.Process.Start("explorer.exe", foldPath1);
        }
        private void CreateExcel(string path)
        {
            DataTable dt = new DataTable();
            clsBuiness.CaiDan_D_PANT cd = new clsBuiness.CaiDan_D_PANT();
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
            //for (int i = 0; i < dataGridView1.Columns.Count; i++)
            //{
            //    //if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("id"))
            //    //{
            //    dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
            //    //}
            //}
            //for (int i = 0; i < dataGridView1.Rows.Count; i++)
            //{
            //    //if (dataGridView1.Rows[i].Cells[6].Value != null)
            //    //{
            //    dt.Rows.Add(dataGridView1.Rows[i].Cells[0].Value, dataGridView1.Rows[i].Cells[1].Value, dataGridView1.Rows[i].Cells[2].Value, dataGridView1.Rows[i].Cells[3].Value, dataGridView1.Rows[i].Cells[4].Value, dataGridView1.Rows[i].Cells[5].Value, dataGridView1.Rows[i].Cells[6].Value, dataGridView1.Rows[i].Cells[7].Value, dataGridView1.Rows[i].Cells[8].Value, dataGridView1.Rows[i].Cells[9].Value, dataGridView1.Rows[i].Cells[10].Value, dataGridView1.Rows[i].Cells[11].Value, dataGridView1.Rows[i].Cells[12].Value, dataGridView1.Rows[i].Cells[13].Value, dataGridView1.Rows[i].Cells[14].Value, dataGridView1.Rows[i].Cells[15].Value, dataGridView1.Rows[i].Cells[16].Value, dataGridView1.Rows[i].Cells[17].Value, dataGridView1.Rows[i].Cells[18].Value, dataGridView1.Rows[i].Cells[19].Value, dataGridView1.Rows[i].Cells[20].Value, dataGridView1.Rows[i].Cells[21].Value, dataGridView1.Rows[i].Cells[22].Value, dataGridView1.Rows[i].Cells[23].Value, dataGridView1.Rows[i].Cells[24].Value, dataGridView1.Rows[i].Cells[25].Value, dataGridView1.Rows[i].Cells[26].Value, dataGridView1.Rows[i].Cells[27].Value, dataGridView1.Rows[i].Cells[28].Value, dataGridView1.Rows[i].Cells[29].Value, dataGridView1.Rows[i].Cells[30].Value, dataGridView1.Rows[i].Cells[31].Value, dataGridView1.Rows[i].Cells[32].Value, dataGridView1.Rows[i].Cells[33].Value, dataGridView1.Rows[i].Cells[34].Value, dataGridView1.Rows[i].Cells[35].Value, dataGridView1.Rows[i].Cells[36].Value, dataGridView1.Rows[i].Cells[37].Value, dataGridView1.Rows[i].Cells[38].Value, dataGridView1.Rows[i].Cells[39].Value, dataGridView1.Rows[i].Cells[40].Value, dataGridView1.Rows[i].Cells[41].Value, dataGridView1.Rows[i].Cells[42].Value);
            //    //}
            //}
            for (int i = 0; i < headerUnitView1.Columns.Count; i++)
            {
                //if (!headerUnitView1.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                //{
                    dt.Columns.Add(headerUnitView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                //}
            }
            for (int i = 0; i < headerUnitView1.Rows.Count; i++)
            {
                if (headerUnitView1.Rows[i].Cells[0].Value != null && headerUnitView1.Rows[i].Cells[0].Value.ToString() != "")
                {
                    dt.Rows.Add(headerUnitView1.Rows[i].Cells[0].Value, headerUnitView1.Rows[i].Cells[1].Value, headerUnitView1.Rows[i].Cells[2].Value, headerUnitView1.Rows[i].Cells[3].Value, headerUnitView1.Rows[i].Cells[4].Value, headerUnitView1.Rows[i].Cells[5].Value, headerUnitView1.Rows[i].Cells[6].Value, headerUnitView1.Rows[i].Cells[7].Value, headerUnitView1.Rows[i].Cells[8].Value, headerUnitView1.Rows[i].Cells[9].Value, headerUnitView1.Rows[i].Cells[10].Value, headerUnitView1.Rows[i].Cells[11].Value, headerUnitView1.Rows[i].Cells[12].Value, headerUnitView1.Rows[i].Cells[13].Value, headerUnitView1.Rows[i].Cells[14].Value, headerUnitView1.Rows[i].Cells[15].Value, headerUnitView1.Rows[i].Cells[16].Value, headerUnitView1.Rows[i].Cells[17].Value, headerUnitView1.Rows[i].Cells[18].Value, headerUnitView1.Rows[i].Cells[19].Value, headerUnitView1.Rows[i].Cells[20].Value, headerUnitView1.Rows[i].Cells[21].Value, headerUnitView1.Rows[i].Cells[22].Value, headerUnitView1.Rows[i].Cells[23].Value, headerUnitView1.Rows[i].Cells[24].Value, headerUnitView1.Rows[i].Cells[25].Value, headerUnitView1.Rows[i].Cells[26].Value, headerUnitView1.Rows[i].Cells[27].Value, headerUnitView1.Rows[i].Cells[28].Value, headerUnitView1.Rows[i].Cells[29].Value, headerUnitView1.Rows[i].Cells[30].Value, headerUnitView1.Rows[i].Cells[31].Value, headerUnitView1.Rows[i].Cells[32].Value, headerUnitView1.Rows[i].Cells[33].Value, headerUnitView1.Rows[i].Cells[34].Value, headerUnitView1.Rows[i].Cells[35].Value, headerUnitView1.Rows[i].Cells[36].Value, headerUnitView1.Rows[i].Cells[37].Value, headerUnitView1.Rows[i].Cells[38].Value, headerUnitView1.Rows[i].Cells[39].Value, headerUnitView1.Rows[i].Cells[40].Value, headerUnitView1.Rows[i].Cells[41].Value, headerUnitView1.Rows[i].Cells[42].Value);
                }
            }
            //cal.InsertChima5(dt, "D.PANT");
            gn2.CDEXCELD_PANT(dt, cd, path);
            foldPath = path + "\\裁单表-" + cd.CaiDanHao + ".xls";
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void frmCaiDanD_PANT_Load(object sender, EventArgs e)
        {
            headerUnitView1.CellValueChanged -= headerUnitView1_CellValueChanged;
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
            txt_zhidan.Text = DateTime.Now.ToString();

            List<clsBuiness.CaiDan_D_PANT> caidan = gn2.selectCaiDanD_PANT("").GroupBy(g => g.CaiDanHao).Select(s => s.First()).ToList<clsBuiness.CaiDan_D_PANT>();
            clsBuiness.CaiDan_D_PANT c = new clsBuiness.CaiDan_D_PANT()
            {
                CaiDanHao = " "
            };
            caidan.Add(c);
            txt_CaidanNo.SelectedIndexChanged -= txt_CaidanNo_SelectedIndexChanged;
            txt_CaidanNo.DataSource = caidan;
            txt_CaidanNo.DisplayMember = "CaiDanHao";
            txt_CaidanNo.ValueMember = "id";
            txt_CaidanNo.SelectedIndexChanged += txt_CaidanNo_SelectedIndexChanged;
            if (txt_CaidanNo.FindString(" ") >= 0)
            {
                txt_CaidanNo.SelectedIndex = txt_CaidanNo.FindString(" ");
            }
            headerUnitView1.CellValueChanged += headerUnitView1_CellValueChanged;
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

        private void headerUnitView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if (e.ColumnIndex >= 6 && e.ColumnIndex <= 40 /*绑定数据源中列的序号*/)
                {
                    //dataGridView1.CellFormatting -= dataGridView1_CellFormatting;
                    double sum = 0;
                    for (int i = 6; i <= 40; i++)
                    {
                        if (headerUnitView1.Rows[e.RowIndex].Cells[i].Value != null && !headerUnitView1.Rows[e.RowIndex].Cells[i].Value.ToString().Equals(string.Empty) && IsNumberic(headerUnitView1.Rows[e.RowIndex].Cells[i].Value.ToString()))
                        {
                            sum = sum + Convert.ToDouble(headerUnitView1.Rows[e.RowIndex].Cells[i].Value);
                        }
                        //dataGridView1.Rows[e.RowIndex].Cells["Column43"].Value = Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column7"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column8"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column9"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column10"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column11"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column12"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column13"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column14"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column15"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column16"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column17"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column18"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column19"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column20"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column21"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column22"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column23"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column24"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column25"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column26"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column27"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column28"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column29"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column30"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column31"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column32"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column33"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column34"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column35"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column36"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column37"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column38"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column39"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column40"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column41"].Value) + Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells["Column42"].Value);
                    }

                    headerUnitView1.Rows[e.RowIndex].Cells["Sub_Total"].Value = sum;
                    //dataGridView1.CellFormatting += dataGridView1_CellFormatting;
                }
            }
        }

        private void txt_CaidanNo_SelectedIndexChanged(object sender, EventArgs e)
        {
            int error = 0;
            try
            {
                headerUnitView1.CellValueChanged -= headerUnitView1_CellValueChanged;

                DataTable dt = new DataTable();
                List<clsBuiness.CaiDan_D_PANT> cdlist = new List<CaiDan_D_PANT>();
                if (!txt_CaidanNo.Text.Equals(string.Empty))
                {
                    cdlist = gn2.selectCaiDanD_PANT(txt_CaidanNo.Text);
                }
                for (int i = 0; i < 43;  i++)
                {
                    //if (!headerUnitView1.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                    //{
                    string tou = headerUnitView1.Columns[i].HeaderText.ToString();
                    if (i == 43)
                    {
                        tou = "id";
                    }
                    dt.Columns.Add(tou, typeof(String));
                    error++;
                    //}
                }
                int foreachi = 0;
                for (int i = 0; i < headerUnitView1.Rows.Count; i++)
                {
                    headerUnitView1.Rows.Clear();
                }
                foreach (CaiDan_D_PANT s in cdlist)
                {
                    headerUnitView1.Rows.Add();
                    headerUnitView1.Rows[foreachi].Cells["LOT"].Value = s.LOT;
                    headerUnitView1.Rows[foreachi].Cells["STYLE"].Value = s.ChimaSTYLE;
                    headerUnitView1.Rows[foreachi].Cells["ART"].Value = s.ART;
                    headerUnitView1.Rows[foreachi].Cells["COLOR"].Value = s.COLOR;
                    headerUnitView1.Rows[foreachi].Cells["COLORName"].Value = s.COLORID;
                    headerUnitView1.Rows[foreachi].Cells["yaowei"].Value = s.JACKET_PANT;
                    headerUnitView1.Rows[foreachi].Cells["C30W_R_30L"].Value = s.C30W_R_30L;
                    headerUnitView1.Rows[foreachi].Cells["C30W_L_32L"].Value = s.C30W_L_32L;
                    headerUnitView1.Rows[foreachi].Cells["C32W_R_30L"].Value = s.C32W_R_30L;
                    headerUnitView1.Rows[foreachi].Cells["C32W_L_32L"].Value = s.C32W_L_32L;
                    headerUnitView1.Rows[foreachi].Cells["C34W_S_38L"].Value = s.C34W_S_28L;
                    headerUnitView1.Rows[foreachi].Cells["C34W_S_39L"].Value = s.C34W_S_29L;
                    headerUnitView1.Rows[foreachi].Cells["C34W_R_30L"].Value = s.C34W_R_30L;
                    headerUnitView1.Rows[foreachi].Cells["C34W_L_32L"].Value = s.C34W_L_32L;
                    headerUnitView1.Rows[foreachi].Cells["C34W_L_34L"].Value = s.C34W_L_34L;
                    headerUnitView1.Rows[foreachi].Cells["C36W_S_28L"].Value = s.C36W_S_28L;
                    headerUnitView1.Rows[foreachi].Cells["C36W_S_29L"].Value = s.C36W_S_29L;
                    headerUnitView1.Rows[foreachi].Cells["C36W_R_30L"].Value = s.C36W_R_30L;
                    headerUnitView1.Rows[foreachi].Cells["C36W_R_31L"].Value = s.C36W_R_31L;
                    headerUnitView1.Rows[foreachi].Cells["C38W_S_28L"].Value = s.C38W_S_28L;
                    headerUnitView1.Rows[foreachi].Cells["C38W_R_30L"].Value = s.C38W_R_30L;
                    headerUnitView1.Rows[foreachi].Cells["C38W_R_31L"].Value = s.C38W_R_31L;
                    headerUnitView1.Rows[foreachi].Cells["C38W_L_32L"].Value = s.C38W_L_32L;
                    headerUnitView1.Rows[foreachi].Cells["C38W_L_34L"].Value = s.C38W_L_34L;
                    headerUnitView1.Rows[foreachi].Cells["C40W_S_28L"].Value = s.C40W_S_28L;
                    headerUnitView1.Rows[foreachi].Cells["C40W_S_29L"].Value = s.C40W_S_29L;
                    headerUnitView1.Rows[foreachi].Cells["C40W_R_30L"].Value = s.C40W_R_30L;
                    headerUnitView1.Rows[foreachi].Cells["C40W_R_31L"].Value = s.C40W_R_31L;
                    headerUnitView1.Rows[foreachi].Cells["C40W_L_32L"].Value = s.C40W_L_32L;
                    headerUnitView1.Rows[foreachi].Cells["C40W_L_34L"].Value = s.C40W_L_34L;
                    headerUnitView1.Rows[foreachi].Cells["C42W_R_30L"].Value = s.C42W_R_30L;
                    headerUnitView1.Rows[foreachi].Cells["C42W_L_32L"].Value = s.C42W_L_32L;
                    headerUnitView1.Rows[foreachi].Cells["C42W_L_34L"].Value = s.C42W_L_34L;
                    headerUnitView1.Rows[foreachi].Cells["C44W_R_30L"].Value = s.C44W_R_30L;
                    headerUnitView1.Rows[foreachi].Cells["C44W_L_32L"].Value = s.C44W_L_32L;
                    headerUnitView1.Rows[foreachi].Cells["C44W_L_34L"].Value = s.C44W_L_34L;
                    headerUnitView1.Rows[foreachi].Cells["C46W_R_30L"].Value = s.C46W_R_30L;
                    headerUnitView1.Rows[foreachi].Cells["C46W_L_32L"].Value = s.C46W_L_32L;
                    headerUnitView1.Rows[foreachi].Cells["C48W_R_30L"].Value = s.C48W_R_30L;
                    headerUnitView1.Rows[foreachi].Cells["C48W_L_32L"].Value = s.C48W_L_32L;
                    headerUnitView1.Rows[foreachi].Cells["C50W_L_32L"].Value = s.C50W_L_32L;
                    headerUnitView1.Rows[foreachi].Cells["Sub_Total"].Value = s.Sub_Total;
                    headerUnitView1.Rows[foreachi].Cells["id"].Value = s.id;
                    foreachi++;
                    //headerUnitView1.Rows[0].Cells["ART"].Value = s.ART;
                    //dt.Rows.Add(s.LOT_, s.STYLE_, s.ART, s.COLOR, s.COLORName, s.yaowei, s.C30W_R_30L, s.C30W_L_32L, s.C32W_R_30L, s.C32W_L_32L, s.C34W_S_38L, s.C34W_S_39L, s.C34W_R_30L, s.C34W_L_32L, s.C34W_L_34L, s.C36W_S_28L, s.C36W_S_29L, s.C36W_R_30L, s.C36W_R_31L, s.C38W_S_28L, s.C38W_R_30L, s.C38W_R_31L, s.C38W_L_32L, s.C38W_L_34L, s.C40W_S_28L, s.C40W_S_29L, s.C40W_R_30L, s.C40W_R_31L, s.C40W_L_32L, s.C40W_L_34L, s.C42W_R_30L, s.C42W_L_32L, s.C42W_L_34L, s.C44W_R_30L, s.C44W_L_32L, s.C44W_L_34L, s.C46W_R_30L, s.C46W_L_32L, s.C48W_R_30L, s.C48W_L_32L, s.C50W_L_32L, s.Sub_Total);
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
                    txt_zhidan.Text = cdlist[0].ZhiDanRiqi.ToString();
                    txt_RN.Text = cdlist[0].RN_NO.ToString();
                    txt_mianlioa.Text = cdlist[0].MianLiao.ToString();
                }
                //headerUnitView1.DataSource = dt;
                headerUnitView1.CellValueChanged += headerUnitView1_CellValueChanged;

            }
            catch (Exception EX) { MessageBox.Show(EX.Message + error); }
        }
        public void ChangeExcel2Image(string filename)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(filename);
            Worksheet sheet = workbook.Worksheets[0];
            imagePath = Directory.GetCurrentDirectory() + "\\image.bmp";

            sheet.SaveToImage(imagePath); //图片后缀.bmp ,imagepath自己设置
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
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                MianFuLiaoDingGou mfldg = new MianFuLiaoDingGou(txt_mianlioa.Text, txt_Style.Text, cb_jgc.Text, txt_CaidanNo.Text, fm);
                if (!mfldg.IsDisposed)
                {
                    if (!HaveOpened(fm, mfldg.Name))
                    {
                        mfldg.MdiParent = fm;
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

        private void headerUnitView1_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["STYLE"].Value = StyleId;
        }

        private void headerUnitView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
