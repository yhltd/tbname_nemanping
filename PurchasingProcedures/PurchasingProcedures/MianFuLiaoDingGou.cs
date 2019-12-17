using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using clsBuiness;
using logic;
using System.Threading;
namespace PurchasingProcedures
{
    public partial class MianFuLiaoDingGou : Form
    {
        protected clsAllnewLogic cal;
        protected GongNeng2 gn;
        protected Define1 df;
        private string ml;
        private string ks;
        private string cdhao;
        private string jgc;
        protected List<HeSuan> hEsuan;
        protected List<HeSuan> mllist;
        private List<clsBuiness.DanHao> cblist;
        private Form fm;
        public MianFuLiaoDingGou(string ml, string ks, string jgc, string cdh,Form f)
        {
            fm = f;
            InitializeComponent();
            cal = new clsAllnewLogic();
            gn = new GongNeng2();
            df = new Define1();

            this.ml = ml;
            this.ks = ks;
            this.jgc = jgc;
            cdhao = cdh;
            if (ml.Equals(string.Empty) || ks.Equals(string.Empty) || jgc.Equals(string.Empty))
            {
                MessageBox.Show("没有找到该面辅料订购表");
                this.Close();
            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void MianFuLiaoDingGou_Load(object sender, EventArgs e)
        {
            try
            {




                //this.comboBox1.SelectedIndexChanged -= comboBox1_SelectedIndexChanged;
                cblist = cal.SelectDanHao("").FindAll(d => d.CaiDanNo.Trim().Equals(cdhao)).GroupBy(gp => gp.Name.Trim()).Select(s => s.First()).ToList<DanHao>();
                List<HeSuan> list1 = GetHeSuanList(ml, "面料");
                DataTable dt1 = DataGirdViewHeader(ml, "面料");

                foreach (HeSuan hs in list1)
                {
                    dt1.Rows.Add(hs.LOT, hs.订单数量, "", hs.色号颜色, hs.单价, hs.预计单耗, hs.预计成本, hs.预计用量, hs.库存, hs.订量, hs.实际到货量, hs.实际到货金额, hs.剩余数量, hs.平均单耗, hs.结算成本);

                }
                this.dataGridView1.DataSource = dt1;
                //this.comboBox1.DataSource = cblist.FindAll(dc => dc.Type.Equals("辅料"));
                //this.comboBox1.ValueMember = "Type";
                //this.comboBox1.DisplayMember = "Name";
                //this.comboBox1.SelectedIndexChanged += comboBox1_SelectedIndexChanged;
                //========================================================

                //string name = comboBox1.SelectedValue.ToString();
                DataTable dt = DataGirdViewHeader2("", "辅料");
                List<HeSuan> list = new List<HeSuan>();
                this.backgroundWorker1.RunWorkerAsync();
                JingDu form = new JingDu(this.backgroundWorker1, "计算中");// 显示进度条窗体
                form.ShowDialog(this);
                form.Close();
                list = CreateFuLiao("", "辅料").GroupBy(g => new { g.Name, g.LOT }).Select(s => s.First()).ToList<clsBuiness.HeSuan>();
                if (list!=null && list.Count <= 0)
                {
                    MessageBox.Show("生成失败！原因：配色表里无该裁单的配色信息！");
                    return;
                }
                foreach (HeSuan hs in list)
                {
                    dt.Rows.Add(hs.Name, hs.LOT, hs.订单数量, "", hs.色号颜色, hs.单价, hs.预计单耗, hs.预计成本, hs.预计用量, hs.库存, hs.订量, hs.实际到货量, hs.实际到货金额, hs.剩余数量, hs.平均单耗, hs.结算成本);

                }
                dataGridView2.Columns.Clear();
                dataGridView2.DataSource = dt;





            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void 裁单输入ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            for (int i = 0; i < 100; i++)
            {
                Thread.Sleep(10);
                worker.ReportProgress(i);
                if (worker.CancellationPending)  // 如果用户取消则跳出处理数据代码 
                {
                    e.Cancel = true;
                    break;
                }
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message);
            }
            else if (e.Cancelled)
            {
            }
            else
            {
            }
        }
        private void ms_caidan_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {




        }

        private List<HeSuan> CreateFuLiao(string name, string type)
        {

            List<HeSuan> endlist = new List<HeSuan>();
            List<clsBuiness.CaiDanALL> cball = new List<clsBuiness.CaiDanALL>();

            List<clsBuiness.CaiDan> caidanlist = gn.selectCaiDan(cdhao);
            List<clsBuiness.CaiDan_RGL2> caidanRGL2 = gn.selectCaiDanRGL2(cdhao);
            List<clsBuiness.CaiDan_SLIM> caidanSLIM = gn.selectCaiDanSLIM(cdhao);
            List<clsBuiness.CaiDan_RGLJ> caidanRGLJ = gn.selectCaiDanRGLJ(cdhao);
            List<clsBuiness.CaiDan_D_PANT> caidanD_PANT = gn.selectCaiDanD_PANT(cdhao);
            List<clsBuiness.CaiDan_C_PANT> caidanC_PANT = gn.selectCaiDanC_PANT(cdhao);
            foreach (clsBuiness.CaiDan cd in caidanlist)
            {
                clsBuiness.CaiDanALL cda = new CaiDanALL();
                cda.Id = cd.Id;
                cda.DESC = cd.DESC;
                cda.FABRIC = cd.FABRIC;
                cda.STYLE = cd.STYLE;
                cda.shuoming = cd.shuoming;
                cda.Jacket = cd.Jacket;
                cda.Pant = cd.Pant;
                cda.LABEL = cd.LABEL;
                cda.JiaGongchang = cd.JiaGongchang;
                cda.CaiDanHao = cd.CaiDanHao;
                cda.ZhiDanRiqi = cd.ZhiDanRiqi;
                cda.JiaoHuoRiqi = cd.JiaoHuoRiqi;
                cda.RN_NO = cd.RN_NO;
                cda.MianLiao = cd.MianLiao;
                cda.LOT = cd.LOT;
                cda.ChimaSTYLE = cd.ChimaSTYLE;
                cda.ART = cd.ART;
                cda.COLOR = cd.COLOR;
                cda.COLORID = cd.COLORID;
                cda.JACKET_PANT = cd.JACKET_PANT;
                cda.C34R = cd.C34R;
                cda.C36R = cd.C36R;
                cda.C38R = cd.C38R;
                cda.C40R = cd.C40R;
                cda.C42R = cd.C42R;
                cda.C44R = cd.C44R;
                cda.C46R = cd.C46R;
                cda.C48R = cd.C48R;
                cda.C50R = cd.C50R;
                cda.C52R = cd.C52R;
                cda.C54R = cd.C54R;
                cda.C56R = cd.C56R;
                cda.C58R = cd.C58R;
                cda.C60R = cd.C60R;
                cda.C62R = cd.C62R;
                cda.C36L = cd.C36L;
                cda.C38L = cd.C38L;
                cda.C40L = cd.C40L;
                cda.C42L = cd.C42L;
                cda.C44L = cd.C44L;
                cda.C46L = cd.C46L;
                cda.C48L = cd.C48L;
                cda.C50L = cd.C50L;
                cda.C52L = cd.C52L;
                cda.C54L = cd.C54L;
                cda.C56L = cd.C56L;
                cda.C58L = cd.C58L;
                cda.C60L = cd.C60L;
                cda.C62L = cd.C62L;
                cda.C34S = cd.C34S;
                cda.C36S = cd.C36S;
                cda.C38S = cd.C38S;
                cda.C40S = cd.C40S;
                cda.C42S = cd.C42S;
                cda.C44S = cd.C44S;
                cda.C46S = cd.C46S;
                cda.Sub_Total = cd.Sub_Total;
                cball.Add(cda);
            }

            foreach (clsBuiness.CaiDan_RGL2 cd in caidanRGL2)
            {
                clsBuiness.CaiDanALL cda = new CaiDanALL();
                cda.Id = cd.Id;
                cda.DESC = cd.DESC;
                cda.FABRIC = cd.FABRIC;
                cda.STYLE = cd.STYLE;
                cda.shuoming = cd.shuoming;
                cda.Jacket = cd.Jacket;
                cda.Pant = cd.Pant;
                cda.LABEL = cd.LABEL;
                cda.JiaGongchang = cd.JiaGongchang;
                cda.CaiDanHao = cd.CaiDanHao;
                cda.ZhiDanRiqi = cd.ZhiDanRiqi;
                cda.JiaoHuoRiqi = cd.JiaoHuoRiqi;
                cda.RN_NO = cd.RN_NO;
                cda.MianLiao = cd.MianLiao;
                cda.LOT = cd.LOT;
                cda.ChimaSTYLE = cd.ChimaSTYLE;
                cda.ART = cd.ART;
                cda.COLOR = cd.COLOR;
                cda.COLORID = cd.COLORID;
                cda.JACKET_PANT = cd.JACKET_PANT;
                cda.C34R = cd.C34R;
                cda.C36R = cd.C36R;
                cda.C38R = cd.C38R;
                cda.C40R = cd.C40R;
                cda.C42R = cd.C42R;
                cda.C44R = cd.C44R;
                cda.C46R = cd.C46R;
                cda.C48R = cd.C48R;
                cda.C50R = cd.C50R;
                cda.C52R = cd.C52R;
                cda.C54R = cd.C54R;
                cda.C56R = cd.C56R;
                cda.C58R = cd.C58R;
                cda.C60R = cd.C60R;
                cda.C62R = cd.C62R;
                cda.C36L = cd.C36L;
                cda.C38L = cd.C38L;
                cda.C40L = cd.C40L;
                cda.C42L = cd.C42L;
                cda.C44L = cd.C44L;
                cda.C46L = cd.C46L;
                cda.C48L = cd.C48L;
                cda.C50L = cd.C50L;
                cda.C52L = cd.C52L;
                cda.C54L = cd.C54L;
                cda.C56L = cd.C56L;
                cda.C58L = cd.C58L;
                cda.C60L = cd.C60L;
                cda.C62L = cd.C62L;
                cda.C34S = cd.C34S;
                cda.C36S = cd.C36S;
                cda.C38S = cd.C38S;
                cda.C40S = cd.C40S;
                cda.C42S = cd.C42S;
                cda.C44S = cd.C44S;
                cda.C46S = cd.C46S;
                cda.Sub_Total = cd.Sub_Total;
                cball.Add(cda);
            }

            foreach (clsBuiness.CaiDan_SLIM cd in caidanSLIM)
            {
                clsBuiness.CaiDanALL cda = new CaiDanALL();
                cda.Id = cd.id;
                cda.DESC = cd.DESC;
                cda.FABRIC = cd.FABRIC;
                cda.STYLE = cd.STYLE;
                cda.shuoming = cd.shuoming;
                cda.Jacket = cd.Jacket;
                cda.Pant = cd.Pant;
                cda.LABEL = cd.LABEL;
                cda.JiaGongchang = cd.JiaGongchang;
                cda.CaiDanHao = cd.CaiDanHao;
                cda.ZhiDanRiqi = cd.ZhiDanRiqi;
                cda.JiaoHuoRiqi = cd.JiaoHuoRiqi;
                cda.RN_NO = cd.RN_NO;
                cda.MianLiao = cd.MianLiao;
                cda.LOT = cd.LOT;
                cda.ChimaSTYLE = cd.ChimaSTYLE;
                cda.ART = cd.ART;
                cda.COLOR = cd.COLOR;
                cda.COLORID = cd.COLORID;
                cda.JACKET_PANT = cd.JACKET_PANT;
                cda.C34R = cd.C34R;
                cda.C36R = cd.C36R;
                cda.C38R = cd.C38R;
                cda.C40R = cd.C40R;
                cda.C42R = cd.C42R;
                cda.C44R = cd.C44R;
                cda.C46R = cd.C46R;
                cda.C48R = cd.C48R;
                cda.C36L = cd.C36L;
                cda.C38L = cd.C38L;
                cda.C40L = cd.C40L;
                cda.C42L = cd.C42L;
                cda.C44L = cd.C44L;
                cda.C46L = cd.C46L;
                cda.C48L = cd.C48L;
                cda.C34S = cd.C34S;
                cda.C36S = cd.C36S;
                cda.C38S = cd.C38S;
                cda.C40S = cd.C40S;
                cda.C42S = cd.C42S;
                cda.C44S = cd.C44S;
                cda.C46S = cd.C46S;
                cda.Sub_Total = cd.Sub_Total;
                cball.Add(cda);
            }

            foreach (clsBuiness.CaiDan_RGLJ cd in caidanRGLJ)
            {
                clsBuiness.CaiDanALL cda = new CaiDanALL();
                cda.Id = cd.id;
                cda.DESC = cd.DESC;
                cda.FABRIC = cd.FABRIC;
                cda.STYLE = cd.STYLE;
                cda.shuoming = cd.shuoming;
                cda.Jacket = cd.Jacket;
                cda.Pant = cd.Pant;
                cda.LABEL = cd.LABEL;
                cda.JiaGongchang = cd.JiaGongchang;
                cda.CaiDanHao = cd.CaiDanHao;
                cda.ZhiDanRiqi = cd.ZhiDanRiqi;
                cda.JiaoHuoRiqi = cd.JiaoHuoRiqi;
                cda.RN_NO = cd.RN_NO;
                cda.MianLiao = cd.MianLiao;
                cda.LOT = cd.LOT;
                cda.ChimaSTYLE = cd.ChimaSTYLE;
                cda.ART = cd.ART;
                cda.COLOR = cd.COLOR;
                cda.COLORID = cd.COLORID;
                cda.JACKET_PANT = cd.JACKET_PANT;
                cda.C34R = cd.C34R;
                cda.C36R = cd.C36R;
                cda.C38R = cd.C38R;
                cda.C40R = cd.C40R;
                cda.C42R = cd.C42R;
                cda.C44R = cd.C44R;
                cda.C46R = cd.C46R;
                cda.C48R = cd.C48R;
                cda.C50R = cd.C50R;
                cda.C52R = cd.C52R;
                cda.C54R = cd.C54R;
                cda.C56R = cd.C56R;
                cda.C58R = cd.C58R;
                cda.C60R = cd.C60R;
                cda.C62R = cd.C62R;
                cda.C36L = cd.C36L;
                cda.C38L = cd.C38L;
                cda.C40L = cd.C40L;
                cda.C42L = cd.C42L;
                cda.C44L = cd.C44L;
                cda.C46L = cd.C46L;
                cda.C48L = cd.C48L;
                cda.C50L = cd.C50L;
                cda.C52L = cd.C52L;
                cda.C54L = cd.C54L;
                cda.C56L = cd.C56L;
                cda.C58L = cd.C58L;
                cda.C60L = cd.C60L;
                cda.C62L = cd.C62L;
                cda.C34S = cd.C34S;
                cda.C36S = cd.C36S;
                cda.C38S = cd.C38S;
                cda.C40S = cd.C40S;
                cda.C42S = cd.C42S;
                cda.C44S = cd.C44S;
                cda.C46S = cd.C46S;
                cda.Sub_Total = cd.Sub_Total;
                cball.Add(cda);
            }

            foreach (clsBuiness.CaiDan_D_PANT cd in caidanD_PANT)
            {
                clsBuiness.CaiDanALL cda = new CaiDanALL();
                cda.Id = cd.id;
                cda.DESC = cd.DESC;
                cda.FABRIC = cd.FABRIC;
                cda.STYLE = cd.STYLE;
                cda.shuoming = cd.shuoming;
                cda.Jacket = cd.Jacket;
                cda.Pant = cd.Pant;
                cda.LABEL = cd.LABEL;
                cda.JiaGongchang = cd.JiaGongchang;
                cda.CaiDanHao = cd.CaiDanHao;
                cda.ZhiDanRiqi = cd.ZhiDanRiqi;
                cda.JiaoHuoRiqi = cd.JiaoHuoRiqi;
                cda.RN_NO = cd.RN_NO;
                cda.MianLiao = cd.MianLiao;
                cda.LOT = cd.LOT;
                cda.ChimaSTYLE = cd.ChimaSTYLE;
                cda.ART = cd.ART;
                cda.COLOR = cd.COLOR;
                cda.COLORID = cd.COLORID;
                cda.JACKET_PANT = cd.JACKET_PANT;
                cda.C30W_R_30L = cd.C30W_R_30L;
                cda.C30W_L_32L = cd.C30W_L_32L;
                cda.C32W_R_30L = cd.C32W_R_30L;
                cda.C32W_L_32L = cd.C32W_L_32L;
                cda.C34W_S_38L = cd.C34W_S_28L;
                cda.C34W_S_39L = cd.C34W_S_29L;
                cda.C34W_R_30L = cd.C34W_R_30L;
                cda.C34W_L_32L = cd.C34W_L_32L;
                cda.C34W_L_34L = cd.C34W_L_34L;
                cda.C36W_S_28L = cd.C36W_S_28L;
                cda.C36W_S_29L = cd.C36W_S_29L;
                cda.C36W_R_30L = cd.C36W_R_30L;
                cda.C36W_R_31L = cd.C36W_R_31L;
                cda.C38W_S_28L = cd.C38W_S_28L;
                cda.C38W_R_30L = cd.C38W_R_30L;
                cda.C38W_R_31L = cd.C38W_R_31L;
                cda.C38W_L_32L = cd.C38W_L_32L;
                cda.C38W_L_34L = cd.C38W_L_34L;
                cda.C40W_S_28L = cd.C40W_S_28L;
                cda.C40W_S_29L = cd.C40W_S_29L;
                cda.C40W_R_30L = cd.C40W_R_30L;
                cda.C40W_R_31L = cd.C40W_R_31L;
                cda.C40W_L_32L = cd.C40W_L_32L;
                cda.C40W_L_34L = cd.C40W_L_34L;
                cda.C42W_R_30L = cd.C42W_R_30L;
                cda.C42W_L_32L = cd.C42W_L_32L;
                cda.C42W_L_34L = cd.C42W_L_34L;
                cda.C44W_R_30L = cd.C44W_R_30L;
                cda.C44W_L_32L = cd.C44W_L_32L;
                cda.C44W_L_34L = cd.C44W_L_34L;
                cda.C46W_R_30L = cd.C46W_R_30L;
                cda.C46W_L_32L = cd.C46W_L_32L;
                cda.C48W_R_30L = cd.C48W_R_30L;
                cda.C48W_L_32L = cd.C48W_L_32L;
                cda.C50W_L_32L = cd.C50W_L_32L;
                cda.Sub_Total = cd.Sub_Total;
                cball.Add(cda);
            }

            foreach (clsBuiness.CaiDan_C_PANT cd in caidanC_PANT)
            {
                clsBuiness.CaiDanALL cda = new CaiDanALL();
                cda.Id = cd.id;
                cda.DESC = cd.DESC;
                cda.FABRIC = cd.FABRIC;
                cda.STYLE = cd.STYLE;
                cda.shuoming = cd.shuoming;
                cda.Jacket = cd.Jacket;
                cda.Pant = cd.Pant;
                cda.LABEL = cd.LABEL;
                cda.JiaGongchang = cd.JiaGongchang;
                cda.CaiDanHao = cd.CaiDanHao;
                cda.ZhiDanRiqi = cd.ZhiDanRiqi;
                cda.JiaoHuoRiqi = cd.JiaoHuoRiqi;
                cda.RN_NO = cd.RN_NO;
                cda.MianLiao = cd.MianLiao;
                cda.LOT = cd.LOT;
                cda.ChimaSTYLE = cd.ChimaSTYLE;
                cda.ART = cd.ART;
                cda.COLOR = cd.COLOR;
                cda.COLORID = cd.COLORID;
                cda.JACKET_PANT = cd.JACKET_PANT;
                cda.C30W_29L = cd.C30W_29L;
                cda.C30W_30L = cd.C30W_30L;
                cda.C30W_32L = cd.C30W_32L;
                cda.C31W_30L = cd.C31W_30L;
                cda.C31W_32L = cd.C31W_32L;
                cda.C32W_28L = cd.C32W_28L;
                cda.C32W_30L = cd.C32W_30L;
                cda.C32W_32L = cd.C32W_32L;
                cda.C33W_29L = cd.C33W_29L;
                cda.C33W_30L = cd.C33W_30L;
                cda.C33W_32L = cd.C33W_32L;
                cda.C33W_34L = cd.C33W_34L;
                cda.C34W_29L = cd.C34W_29L;
                cda.C34W_30L = cd.C34W_30L;
                cda.C34W_31L = cd.C34W_31L;
                cda.C34W_32L = cd.C34W_32L;
                cda.C34W_34L = cd.C34W_34L;
                cda.C36W_29L = cd.C36W_29L;
                cda.C36W_30L = cd.C36W_30L;
                cda.C36W_32L = cd.C36W_32L;
                cda.C36W_34L = cd.C36W_34L;
                cda.C38W_29L = cd.C38W_29L;
                cda.C38W_30L = cd.C38W_30L;
                cda.C38W_32L = cd.C38W_32L;
                cda.C38W_34L = cd.C38W_34L;
                cda.C40W_28L = cd.C40W_28L;
                cda.C40W_30L = cd.C40W_30L;
                cda.C40W_32L = cd.C40W_32L;
                cda.C40W_34L = cd.C40W_34L;
                cda.C42W_30L = cd.C42W_30L;
                cda.C42W_32L = cd.C42W_32L;
                cda.C42W_34L = cd.C42W_34L;
                cda.C44W_29L = cd.C44W_29L;
                cda.C44W_30L = cd.C44W_30L;
                cda.C44W_32L = cd.C44W_32L;
                cda.Sub_Total = cd.Sub_Total;
                cball.Add(cda);
            }

            //List<clsBuiness.CaiDan> caidanlist = gn.selectCaiDan(cdhao);

            //List<clsBuiness.CaiDan> caidanlist = gn.selectCaiDan(cdhao);

            //List<clsBuiness.CaiDan> caidanlist = gn.selectCaiDan(cdhao);
            //List<clsBuiness.CaiDan> caidanlist = gn.selectCaiDan(cdhao);


            if (cball.Count == 0) 
            {
            


            }
            List<HeSuan> hs = new List<HeSuan>();
            List<DanHao> dh = cblist.FindAll(dc => dc.Type.Equals("辅料"));
            Dictionary<string, string> dic = new Dictionary<string, string>();
            List<PeiSe> peis = cal.selectPeise("");
            List<PeiSe> peisa = cal.selectPeise("");
            List<clsBuiness.GongHuoFang> ghflist = df.selectGongHuoFang();
            List<clsBuiness.DanHao> dhlist = cal.SelectDanHao("");
            List<clsBuiness.KuCun> kclist = cal.SelectKC();
            int errorlog = 0;
            try
            {
                //List<PeiSe> pslist = new List<PeiSe>();
                //foreach (clsBuiness.DanHao dca in dh) 
                //{

                //}

                //IEnumerable<int> en = peis.Intersect(peisa);
                //foreach(clsBuiness.DanHao dcc in dh){
                //List<PeiSe> peiselist = peis.FindAll(p => p.PingMing.Equals(dcc.Name));
                List<PeiSe> peiselist = cal.selectps();
                //获取配色表里的信息 加以计算（色号+总数）start

                List<string> key = new List<string>();
                //dic.Keys = null;
                foreach (PeiSe ps in peiselist)
                {
                    name = ps.PingMing;
                    if (cball.FindAll(f => f.ART.Equals(ps.HuoHao)).Count > 0)
                    {
                        foreach (clsBuiness.CaiDanALL cd in cball)
                        {
                            if (cd.ART.Equals(ps.HuoHao))
                            {
                                if (dic.ContainsKey(ps.PingMing + "=" + ps.HuoHao))
                                {
                                    int dicJia = 0;
                                    int jia2 = 0;
                                    if (!dic[ps.PingMing + "=" + ps.HuoHao].Split('=')[0].Equals(string.Empty)) 
                                    {
                                        dicJia = Convert.ToInt32(dic[ps.PingMing + "=" + ps.HuoHao].Split('=')[0]);
                                    }
                                    if (!cd.Sub_Total.Equals(string.Empty)) 
                                    {
                                        jia2 = Convert.ToInt32(cd.Sub_Total);
                                    }
                                    dic[ps.PingMing + "=" + ps.HuoHao] = (dicJia + jia2).ToString();
                                    if (dic[ps.PingMing + "=" + ps.HuoHao].Split('=').Length < 2)
                                    {
                                        dic[ps.PingMing + "=" + ps.HuoHao] = dic[ps.PingMing + "=" + ps.HuoHao] + "=" + cd.COLORID.Trim() + " " + cd.COLOR.Trim(); //色号&颜色
                                    }
                                    else
                                    {
                                        dic[ps.PingMing + "=" + ps.HuoHao] = dic[ps.PingMing + "=" + ps.HuoHao] + "=" + " ";
                                    }
                                }
                                else
                                {
                                    dic.Add(ps.PingMing + "=" + ps.HuoHao, cd.Sub_Total+ "=" + cd.COLORID.Trim() + " " + cd.COLOR.Trim());
                                }
                            }

                        }
                    }

                    //end 获取配色表里的信息 加以计算（色号+总数）
                    //foreach (clsBuiness.CaiDan c in caidanlist)
                    //{
                    //    if (dic.ContainsKey(dcc.Name + "=" + ps.HuoHao))
                    //    {
                    //        if (dic[dcc.Name + "=" + ps.HuoHao].Split('=').Length < 2)
                    //        {
                    //            dic[dcc.Name + "=" + ps.HuoHao] = dic[dcc.Name + "=" + ps.HuoHao] + "=" + c.COLORID.Trim() + " " + c.COLOR.Trim(); //色号&颜色
                    //        }
                    //        else 
                    //        {
                    //            dic[dcc.Name + "=" + ps.HuoHao] = dic[dcc.Name + "=" + ps.HuoHao] + "=" + " ";
                    //        }
                    //    }
                    //}

                    key = new List<string>(dic.Keys);
                    
                    errorlog++;
                    for (int i = 0; i < key.Count; i++)
                    {
                        foreach (clsBuiness.GongHuoFang g in ghflist)
                        {
                            if (dic.ContainsKey(key[i]))
                            {
                                if (dic[key[i]].Split('=')[1].Equals(g.SeHao.Trim() + g.Yanse.Trim()) && g.PingMing.Trim().Equals(name))
                                {
                                    dic[key[i]] = dic[key[i]] + "=" + g.DanJia;//单价
                                }
                                else
                                {
                                    dic[key[i]] = dic[key[i]] + "=" + "";
                                }

                            }
                        }
                        foreach (clsBuiness.DanHao d in dhlist)
                        {
                            if (dic.ContainsKey(key[i]))
                            {
                                if (d.Style.Equals(ks) && d.JiaGongChang.Equals(jgc) && d.Name.Trim().Equals(name) && d.Yanse.Trim().Equals(dic[key[i]].Split('=')[1].Trim() + " " + dic[key[i]].Split('=')[2].Trim()))
                                {
                                    dic[key[i]] = dic[key[i]] + "=" + d.DanHao1;//预计单耗
                                    dic[key[i]] = dic[key[i]] + "=" + Convert.ToInt32(dic[key[i]].Split('=')[2]) * Convert.ToInt32(d.DanHao1);//预计成本
                                    dic[key[i]] = dic[key[i]] + "=" + Convert.ToInt32(dic[key[i]].Split('=')[0]) * Convert.ToInt32(d.DanHao1);//预计用量
                                }
                                else
                                {
                                    dic[key[i]] = dic[key[i]] + "=" + "";
                                }
                            }
                        }
                      


                        try
                        {
                            foreach (clsBuiness.KuCun kc in kclist)
                            {
                               


                                if (dic.ContainsKey(key[i]))
                                {
                                    if (kc.PingMing.Trim().Equals(name) && kc.SeHao.Trim().Equals(dic[key[i]].Split('=')[1]))
                                    {
                                        dic[key[i]] = dic[key[i]] + "=" + kc.ShuLiang;
                                        dic[key[i]] = dic[key[i]] + "=" + (Convert.ToInt32(dic[key[i]].Split('=')[5].ToString()) - Convert.ToInt32(kc.ShuLiang.ToString())); //订量
                                        dic[key[i]] = dic[key[i]] + "=" + " ";
                                        //dic[kvp.Key] = dic[kvp.Key] + "=" +
                                    }
                                    else
                                    {
                                        dic[key[i]] = dic[key[i]] + "=" + "";
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {

                            throw;
                        }

                    }

                    for (int i = 0; i < key.Count; i++)
                    {
                        HeSuan endhs = new HeSuan();

                        if (key[i].Split('=').Length > 0)
                        {
                            endhs.Name = key[i].Split('=')[0];
                        }
                        if (key[i].Split('=').Length > 1)
                        {
                            endhs.LOT = key[i].Split('=')[1];
                        }

                        if (dic[key[i]].Split('=').Length > 0)
                        {
                            endhs.订单数量 = dic[key[i]].Split('=')[0];
                        }

                        if (dic[key[i]].Split('=').Length > 1)
                        {
                            endhs.色号颜色 = dic[key[i]].Split('=')[1];
                        }
                        if (dic[key[i]].Split('=').Length > 2)
                        {
                            endhs.单价 = dic[key[i]].Split('=')[2];
                        }
                        if (dic[key[i]].Split('=').Length > 3)
                        {
                            endhs.预计单耗 = dic[key[i]].Split('=')[3];
                        }
                        if (dic[key[i]].Split('=').Length > 4)
                        {
                            endhs.预计成本 = dic[key[i]].Split('=')[4];
                        }
                        if (dic[key[i]].Split('=').Length > 5)
                        {
                            endhs.预计用量 = dic[key[i]].Split('=')[5];
                        }
                        if (dic[key[i]].Split('=').Length > 6)
                        {
                            endhs.库存 = dic[key[i]].Split('=')[6];
                        }
                        if (dic[key[i]].Split('=').Length > 7)
                        {
                            endhs.订量 = dic[key[i]].Split('=')[7];
                        }
                        if (dic[key[i]].Split('=').Length > 8)
                        {
                            endhs.实际到货量 = dic[key[i]].Split('=')[8];
                        }

                        if (endhs.实际到货量 != null && endhs.单价 != null && !endhs.实际到货量.Equals(string.Empty) && !endhs.单价.Equals(string.Empty))
                        {
                            endhs.实际到货金额 = (Convert.ToInt32(endhs.单价) * Convert.ToInt32(endhs.实际到货量)).ToString();
                        }
                        //if (!endhs.库存.Equals(string.Empty) && !endhs.实际到货量.Equals(string.Empty) && !endhs.剩余数量.Equals(string.Empty))
                        //{
                        //   endhs.平均单耗 = (endhs.单价-endhs.剩余数量+endhs.库存)
                        //}
                        endlist.Add(endhs);

                    }



                }
            }
            catch (Exception ex)
            {

                throw;
            }

            //}
            hEsuan = endlist;
            return endlist;
        }
        private static DataTable DataGirdViewHeader2(string name, string type)
        {


            DataTable dt = new DataTable();
            dt.Columns.Add("辅料名称", typeof(string));
            dt.Columns.Add("货号", typeof(string));
            dt.Columns.Add("订单数量", typeof(string));
            dt.Columns.Add("实际出口数量", typeof(string));
            dt.Columns.Add("色号&颜色", typeof(string));
            dt.Columns.Add("单价", typeof(string));
            dt.Columns.Add(name + "预计单耗", typeof(string));
            dt.Columns.Add(name + "预计成本", typeof(string));
            dt.Columns.Add(name + "预计用量", typeof(string));
            dt.Columns.Add(name + "库存", typeof(string));
            dt.Columns.Add(name + "订量", typeof(string));
            dt.Columns.Add(name + "实际到货量", typeof(string));
            dt.Columns.Add(name + "实际到货金额", typeof(string));
            dt.Columns.Add(name + "剩余数量", typeof(string));
            dt.Columns.Add(name + "平均单耗", typeof(string));
            dt.Columns.Add(name + "结算成本", typeof(string));
            if (type.Equals("辅料"))
            {
                dt.Columns.Add("小计", typeof(string));
            }
            else
            {
                dt.Columns.Add("总数", typeof(string));
            }
            return dt;
        }

        private static DataTable DataGirdViewHeader(string name, string type)
        {


            DataTable dt = new DataTable();
            dt.Columns.Add("LOT#", typeof(string));
            dt.Columns.Add("订单数量", typeof(string));
            dt.Columns.Add("实际出口数量", typeof(string));
            dt.Columns.Add("色号&颜色", typeof(string));
            dt.Columns.Add("单价", typeof(string));
            dt.Columns.Add(name + "预计单耗", typeof(string));
            dt.Columns.Add(name + "预计成本", typeof(string));
            dt.Columns.Add(name + "预计用量", typeof(string));
            dt.Columns.Add(name + "库存", typeof(string));
            dt.Columns.Add(name + "订量", typeof(string));
            dt.Columns.Add(name + "实际到货量", typeof(string));
            dt.Columns.Add(name + "实际到货金额", typeof(string));
            dt.Columns.Add(name + "剩余数量", typeof(string));
            dt.Columns.Add(name + "平均单耗", typeof(string));
            dt.Columns.Add(name + "结算成本", typeof(string));
            if (type.Equals("辅料"))
            {
                dt.Columns.Add("小计", typeof(string));
            }
            else
            {
                dt.Columns.Add("总数", typeof(string));
            }
            return dt;
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
        private List<HeSuan> GetHeSuanList(string name, string type)
        {
            List<clsBuiness.CaiDan> caidanlist = gn.selectCaiDan(cdhao);
            List<clsBuiness.CaiDan_RGL2> caidanlistRGL2 = gn.selectCaiDanRGL2(cdhao);
            List<clsBuiness.CaiDan_SLIM> caidanlistSLIM = gn.selectCaiDanSLIM(cdhao);
            List<clsBuiness.CaiDan_RGLJ> caidanlistRGLJ = gn.selectCaiDanRGLJ(cdhao);
            List<clsBuiness.CaiDan_D_PANT> caidanlistD_PANT = gn.selectCaiDanD_PANT(cdhao);
            List<HeSuan> hs = new List<HeSuan>();
            Dictionary<string, string> dic = new Dictionary<string, string>();
            foreach (clsBuiness.CaiDan c in caidanlist)//成衣数量
            {
                if (dic.ContainsKey(c.LOT))
                {
                    if (c.Sub_Total != null && c.Sub_Total != "" && !IsNumberic(c.Sub_Total) && dic[c.LOT] != null && dic[c.LOT] != "" && !IsNumberic(dic[c.LOT]))
                    {


                        dic[c.LOT] = (Convert.ToInt32(dic[c.LOT]) + Convert.ToInt32(c.Sub_Total)).ToString();
                    }
                }
                else
                {
                    dic.Add(c.LOT, c.Sub_Total);
                }

            }
            foreach (clsBuiness.CaiDan c in caidanlist)
            {
                if (dic[c.LOT].Split('=').Length < 2)
                {

                    dic[c.LOT] = dic[c.LOT] + "=" + c.COLORID.Trim() + c.COLOR.Trim(); //色号&颜色
                }
            }
            //YAO
            foreach (clsBuiness.CaiDan_RGL2 c in caidanlistRGL2)//成衣数量
            {
                if (dic.ContainsKey(c.LOT))
                {
                    if (c.Sub_Total != null && c.Sub_Total != "" && !IsNumberic(c.Sub_Total) && dic[c.LOT] != null && dic[c.LOT] != "" && !IsNumberic(dic[c.LOT]))
                    {
                        dic[c.LOT] = (Convert.ToInt32(dic[c.LOT]) + Convert.ToInt32(c.Sub_Total)).ToString();
                    }
                }
                else
                {
                    dic.Add(c.LOT, c.Sub_Total);
                }

            }
            foreach (clsBuiness.CaiDan_RGL2 c in caidanlistRGL2)
            {
                if (dic[c.LOT].Split('=').Length < 2)
                {
                   
                    dic[c.LOT] = dic[c.LOT] + "=" + c.COLORID.Trim() + c.COLOR.Trim(); //色号&颜色
                }
            }



            foreach (clsBuiness.CaiDan_SLIM c in caidanlistSLIM)//成衣数量
            {
                if (dic.ContainsKey(c.LOT))
                {
                    if (c.Sub_Total != null && c.Sub_Total != "" && !IsNumberic(c.Sub_Total) && dic[c.LOT] != null && dic[c.LOT] != "" && !IsNumberic(dic[c.LOT]))
                    {
                        dic[c.LOT] = (Convert.ToInt32(dic[c.LOT]) + Convert.ToInt32(c.Sub_Total)).ToString();
                    }
                }
                else
                {
                    dic.Add(c.LOT, c.Sub_Total);
                }

            }
            foreach (clsBuiness.CaiDan_SLIM c in caidanlistSLIM)
            {
                if (dic[c.LOT].Split('=').Length < 2)
                {
                    dic[c.LOT] = dic[c.LOT] + "=" + c.COLORID.Trim() + c.COLOR.Trim(); //色号&颜色
                }
            }


            foreach (clsBuiness.CaiDan_RGLJ c in caidanlistRGLJ)//成衣数量
            {

                if (dic.ContainsKey(c.LOT))
                {
                    if (c.Sub_Total != null && c.Sub_Total != "" && !IsNumberic(c.Sub_Total) && dic[c.LOT] != null && dic[c.LOT] != "" && !IsNumberic(dic[c.LOT]))
                    {
                        dic[c.LOT] = (Convert.ToInt32(dic[c.LOT]) + Convert.ToInt32(c.Sub_Total)).ToString();
                    }
                }
                else
                {
                    dic.Add(c.LOT, c.Sub_Total);
                }

            }
            foreach (clsBuiness.CaiDan_RGLJ c in caidanlistRGLJ)
            {
                if (dic[c.LOT].Split('=').Length < 2)
                {
                    dic[c.LOT] = dic[c.LOT] + "=" + c.COLORID.Trim() + c.COLOR.Trim(); //色号&颜色
                }
            }




            foreach (clsBuiness.CaiDan_D_PANT c in caidanlistD_PANT)//成衣数量
            {
                if (dic.ContainsKey(c.LOT))
                {
                    if (c.Sub_Total != null && c.Sub_Total != "" && !IsNumberic(c.Sub_Total) && dic[c.LOT] != null && dic[c.LOT] != "" && !IsNumberic(dic[c.LOT]))
                    {
                        dic[c.LOT] = (Convert.ToInt32(dic[c.LOT]) + Convert.ToInt32(c.Sub_Total)).ToString();
                    }
                }
                else
                {
                    dic.Add(c.LOT, c.Sub_Total);
                }

            }
            foreach (clsBuiness.CaiDan_D_PANT c in caidanlistD_PANT)
            {
                if (dic[c.LOT].Split('=').Length < 2)
                {
                    dic[c.LOT] = dic[c.LOT] + "=" + c.COLORID.Trim() + c.COLOR.Trim(); //色号&颜色
                }
            }



            List<clsBuiness.GongHuoFang> ghflist = df.selectGongHuoFang();
            // foreach (KeyValuePair<string,string>kvp in dic)
            //{               

            //}
            List<clsBuiness.DanHao> dhlist = cal.SelectDanHao("");
            List<clsBuiness.KuCun> kclist = cal.SelectKC();
            List<string> key = new List<string>(dic.Keys);

            for (int i = 0; i < key.Count; i++)
            {
                foreach (clsBuiness.GongHuoFang g in ghflist)
                {
                    if (dic.ContainsKey(key[i]))
                    {
                        if (dic[key[i]].Split('=')[1].Equals(g.SeHao.Trim() + g.Yanse.Trim()) && g.PingMing.Trim().Equals(name))
                        {
                            dic[key[i]] = dic[key[i]] + "=" + g.DanJia;//单价
                        }
                        else
                        {
                            dic[key[i]] = dic[key[i]] + "=" + "";
                        }
                    }
                }
                foreach (clsBuiness.DanHao d in dhlist)
                {
                    if (dic.ContainsKey(key[i]))
                    {
                        if (d.Style.Equals(ks) && d.JiaGongChang.Equals(jgc) && d.Name.Trim().Equals(name) && d.Yanse.Trim().Equals(dic[key[i]].Split('=')[1].Trim() + " " + dic[key[i]].Split('=')[2].Trim()))
                        {
                            dic[key[i]] = dic[key[i]] + "=" + d.DanHao1;//预计单耗
                            dic[key[i]] = dic[key[i]] + "=" + Convert.ToInt32(dic[key[i]].Split('=')[2]) * Convert.ToInt32(d.DanHao1);//预计成本
                            dic[key[i]] = dic[key[i]] + "=" + Convert.ToInt32(dic[key[i]].Split('=')[0]) * Convert.ToInt32(d.DanHao1);//预计用量
                        }
                        else
                        {
                            dic[key[i]] = dic[key[i]] + "=" + "";
                        }
                    }
                }
                foreach (clsBuiness.KuCun kc in kclist)
                {
                    if (dic.ContainsKey(key[i]))
                    {
                        if (kc.PingMing.Trim().Equals(name) && kc.SeHao.Trim().Equals(dic[key[i]].Split('=')[1]))
                        {
                            dic[key[i]] = dic[key[i]] + "=" + kc.ShuLiang;
                            dic[key[i]] = dic[key[i]] + "=" + (Convert.ToInt32(dic[key[i]].Split('=')[5].ToString()) - Convert.ToInt32(kc.ShuLiang.ToString())); //订量
                            dic[key[i]] = dic[key[i]] + "=" + " ";
                            //dic[kvp.Key] = dic[kvp.Key] + "=" +
                        }
                        else
                        {
                            dic[key[i]] = dic[key[i]] + "=" + "";
                        }
                    }
                }

            }
            List<HeSuan> list = new List<HeSuan>();
            for (int i = 0; i < key.Count; i++)
            {
                HeSuan endhs = new HeSuan();
                endhs.LOT = key[i];
                if (dic[key[i]].Split('=').Length > 0)
                {
                    endhs.订单数量 = dic[key[i]].Split('=')[0];
                }
                endhs.Name = key[i];
                if (dic[key[i]].Split('=').Length > 1)
                {
                    endhs.色号颜色 = dic[key[i]].Split('=')[1];
                }
                if (dic[key[i]].Split('=').Length > 2)
                {
                    endhs.单价 = dic[key[i]].Split('=')[2];
                }
                if (dic[key[i]].Split('=').Length > 3)
                {
                    endhs.预计单耗 = dic[key[i]].Split('=')[3];
                }
                if (dic[key[i]].Split('=').Length > 4)
                {
                    endhs.预计成本 = dic[key[i]].Split('=')[4];
                }
                if (dic[key[i]].Split('=').Length > 5)
                {
                    endhs.预计用量 = dic[key[i]].Split('=')[5];
                }
                if (dic[key[i]].Split('=').Length > 6)
                {
                    endhs.库存 = dic[key[i]].Split('=')[6];
                }
                if (dic[key[i]].Split('=').Length > 7)
                {
                    endhs.订量 = dic[key[i]].Split('=')[7];
                }
                if (dic[key[i]].Split('=').Length > 8)
                {
                    endhs.实际到货量 = dic[key[i]].Split('=')[8];
                }

                if (endhs.实际到货量 != null && endhs.单价 != null && !endhs.实际到货量.Equals(string.Empty) && !endhs.单价.Equals(string.Empty))
                {
                    endhs.实际到货金额 = (Convert.ToInt32(endhs.单价) * Convert.ToInt32(endhs.实际到货量)).ToString();
                }
                //if (!endhs.库存.Equals(string.Empty) && !endhs.实际到货量.Equals(string.Empty) && !endhs.剩余数量.Equals(string.Empty))
                //{
                //   endhs.平均单耗 = (endhs.单价-endhs.剩余数量+endhs.库存)
                //}
                list.Add(endhs);
            }
            mllist = list;
            return list;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.backgroundWorker1.RunWorkerAsync();
            JingDu form = new JingDu(this.backgroundWorker1, "生成中");// 显示进度条窗体
            form.ShowDialog(this);
            form.Close();
            List<HeSuan> FL = new List<HeSuan>();
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                if (dataGridView2[0, i].Value != null)
                {
                    HeSuan h = new HeSuan()
                    {
                        Name = dataGridView2[0, i].Value.ToString(),
                        LOT = dataGridView2[1, i].Value.ToString(),
                        订单数量 = dataGridView2[2, i].Value.ToString(),

                        色号颜色 = dataGridView2[4, i].Value.ToString(),
                        单价 = dataGridView2[5, i].Value.ToString(),
                        预计单耗 = dataGridView2[6, i].Value.ToString(),
                        预计成本 = dataGridView2[7, i].Value.ToString(),
                        预计用量 = dataGridView2[8, i].Value.ToString(),
                        库存 = dataGridView2[9, i].Value.ToString(),
                        订量 = dataGridView2[10, i].Value.ToString(),
                        实际到货量 = dataGridView2[11, i].Value.ToString(),
                        实际到货金额 = dataGridView2[12, i].Value.ToString(),
                        剩余数量 = dataGridView2[13, i].Value.ToString(),
                        平均单耗 = dataGridView2[14, i].Value.ToString(),
                        结算成本 = dataGridView2[15, i].Value.ToString(),
                    };
                    FL.Add(h);
                }
            }

            if (FL.Count > 0)
            {
                mflDgd mfl = new mflDgd(hEsuan, cdhao, jgc, ks);
                mfl.Show();

            }
            else
            {
                MessageBox.Show("生成失败！ 原因：辅料为空！");
                return;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                this.backgroundWorker1.RunWorkerAsync();
                JingDu form = new JingDu(this.backgroundWorker1, "生成中");// 显示进度条窗体
                form.ShowDialog(this);
                form.Close();
                List<HeSuan> ML = new List<HeSuan>();
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1[0, i].Value != null)
                    {
                        HeSuan h = new HeSuan()
                        {
                            LOT = dataGridView1[0, i].Value.ToString(),
                            订单数量 = dataGridView1[1, i].Value.ToString(),
                            //Name = "面料", 
                            色号颜色 = dataGridView1[3, i].Value.ToString(),
                            单价 = dataGridView1[4, i].Value.ToString(),
                            预计单耗 = dataGridView1[5, i].Value.ToString(),
                            预计成本 = dataGridView1[6, i].Value.ToString(),
                            预计用量 = dataGridView1[7, i].Value.ToString(),
                            库存 = dataGridView1[8, i].Value.ToString(),
                            订量 = dataGridView1[9, i].Value.ToString(),
                            实际到货量 = dataGridView1[10, i].Value.ToString(),
                            实际到货金额 = dataGridView1[11, i].Value.ToString(),
                            剩余数量 = dataGridView1[12, i].Value.ToString(),
                            平均单耗 = dataGridView1[13, i].Value.ToString(),
                            结算成本 = dataGridView1[14, i].Value.ToString(),
                        };
                        ML.Add(h);
                    }
                }
                List<HeSuan> FL = new List<HeSuan>();
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    if (dataGridView2[0, i].Value != null)
                    {
                        HeSuan h = new HeSuan()
                        {
                            Name = dataGridView2[0, i].Value.ToString(),
                            LOT = dataGridView2[1, i].Value.ToString(),
                            订单数量 = dataGridView2[2, i].Value.ToString(),

                            色号颜色 = dataGridView2[4, i].Value.ToString(),
                            单价 = dataGridView2[5, i].Value.ToString(),
                            预计单耗 = dataGridView2[6, i].Value.ToString(),
                            预计成本 = dataGridView2[7, i].Value.ToString(),
                            预计用量 = dataGridView2[8, i].Value.ToString(),
                            库存 = dataGridView2[9, i].Value.ToString(),
                            订量 = dataGridView2[10, i].Value.ToString(),
                            实际到货量 = dataGridView2[11, i].Value.ToString(),
                            实际到货金额 = dataGridView2[12, i].Value.ToString(),
                            剩余数量 = dataGridView2[13, i].Value.ToString(),
                            平均单耗 = dataGridView2[14, i].Value.ToString(),
                            结算成本 = dataGridView2[15, i].Value.ToString(),
                        };
                        FL.Add(h);
                    }
                }

                shengchengBiaoge scbg = new shengchengBiaoge(cdhao, ML, FL);
                scbg.Show();
            }
            catch (Exception ex)
            {

            }
        }

        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 5 || e.ColumnIndex == 6)
            {
                if (!dataGridView2.Rows[e.RowIndex].Cells[5].Value.ToString().Equals(string.Empty) && !dataGridView2.Rows[e.RowIndex].Cells[6].Value.ToString().Equals(" ") && !dataGridView2.Rows[e.RowIndex].Cells[6].Value.ToString().Equals(string.Empty))
                {
                    dataGridView2.Rows[e.RowIndex].Cells[7].Value = Convert.ToDouble(dataGridView2.Rows[e.RowIndex].Cells[5].Value.ToString()) * Convert.ToDouble(dataGridView2.Rows[e.RowIndex].Cells[6].Value);
                }
            }
            else if (e.ColumnIndex == 6 || e.ColumnIndex == 2)
            {
                if (!dataGridView2.Rows[e.RowIndex].Cells[6].Value.ToString().Trim().Equals(string.Empty) && !dataGridView2.Rows[e.RowIndex].Cells[2].Value.ToString().Trim().Equals(string.Empty))
                {
                    dataGridView2.Rows[e.RowIndex].Cells[8].Value = Convert.ToDouble(dataGridView2.Rows[e.RowIndex].Cells[2].Value.ToString()) * Convert.ToDouble(dataGridView2.Rows[e.RowIndex].Cells[6].Value.ToString());
                }
            }
            else if (e.ColumnIndex == 8 || e.ColumnIndex == 9)
            {
                if (!dataGridView2.Rows[e.RowIndex].Cells[8].Value.ToString().Trim().Equals(string.Empty) && !dataGridView2.Rows[e.RowIndex].Cells[9].Value.ToString().Trim().Equals(string.Empty))
                {
                    dataGridView2.Rows[e.RowIndex].Cells[10].Value = Convert.ToDouble(dataGridView2.Rows[e.RowIndex].Cells[8].Value.ToString()) - Convert.ToDouble(dataGridView2.Rows[e.RowIndex].Cells[9].Value);
                }
            }
            else if (e.ColumnIndex == 11 || e.ColumnIndex == 5)
            {
                if (!dataGridView2.Rows[e.RowIndex].Cells[11].Value.ToString().Equals(string.Empty) && !dataGridView2.Rows[e.RowIndex].Cells[5].Value.ToString().Equals(string.Empty))
                {
                    dataGridView2.Rows[e.RowIndex].Cells[12].Value = Convert.ToDouble(dataGridView2.Rows[e.RowIndex].Cells[5].Value.ToString()) * Convert.ToDouble(dataGridView2.Rows[e.RowIndex].Cells[11].Value);
                }
            }
            else if (e.ColumnIndex == 11 || e.ColumnIndex == 9)
            {
                if (!dataGridView2.Rows[e.RowIndex].Cells[11].Value.ToString().Equals(string.Empty) && !dataGridView2.Rows[e.RowIndex].Cells[9].Value.ToString().Equals(string.Empty))
                {
                    dataGridView2.Rows[e.RowIndex].Cells[14].Value = (Convert.ToDouble(dataGridView2.Rows[e.RowIndex].Cells[11].Value.ToString()) + Convert.ToDouble(dataGridView2.Rows[e.RowIndex].Cells[9].Value) - Convert.ToDouble(dataGridView2.Rows[e.RowIndex].Cells[13].Value)) / Convert.ToDouble(dataGridView2.Rows[e.RowIndex].Cells[3].Value);
                }
            }
            if (e.ColumnIndex == 12)
            {
                Dictionary<string, string> dic = new Dictionary<string, string>();
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    if (dataGridView2.Rows[i].Cells[12].Value != null && !dataGridView2.Rows[i].Cells[12].Value.ToString().Equals(string.Empty))
                    {
                        if (dataGridView2.Rows[i].Cells[0].Value != null)
                        {

                            string keys = dataGridView2.Rows[i].Cells[0].Value.ToString().Trim();
                            if (dic.ContainsKey(keys))
                            {
                                dic[keys] = (Convert.ToInt32(dic[keys].Split('=')[0]) + Convert.ToInt32(dataGridView2.Rows[i].Cells[12].Value.ToString())).ToString() + "=" + dic[keys].Split('=')[1];
                            }
                            else
                            {
                                dic.Add(keys, dataGridView2.Rows[i].Cells[12].Value.ToString() + "=" + i);
                            }
                        }
                    }
                }
                List<string> key = new List<string>(dic.Keys);
                double chushu = 0;
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1.Rows[i].Cells[2].Value != null)
                    {
                        chushu = chushu + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value.ToString());
                    }
                }
                foreach (string k in key)
                {
                    dataGridView2.Rows[Convert.ToInt32(dic[k].Split('=')[1])].Cells[15].Value = Convert.ToDouble(dic[k].Split('=')[0]) / chushu;
                }
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 4 || e.ColumnIndex == 5)
            {
                if (!dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString().Equals(string.Empty) && !dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString().Equals(" ") && !dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString().Equals(string.Empty))
                {
                    dataGridView1.Rows[e.RowIndex].Cells[6].Value = Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString()) * Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[4].Value);
                }
            }
            else if (e.ColumnIndex == 5 || e.ColumnIndex == 1)
            {
                if (!dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString().Equals(string.Empty) && !dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString().Equals(string.Empty))
                {
                    dataGridView1.Rows[e.RowIndex].Cells[7].Value = Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString()) * Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString());
                }
            }
            else if (e.ColumnIndex == 7 || e.ColumnIndex == 8)
            {
                if (!dataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString().Equals(string.Empty) && !dataGridView1.Rows[e.RowIndex].Cells[8].Value.ToString().Equals(string.Empty))
                {
                    dataGridView1.Rows[e.RowIndex].Cells[9].Value = Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString()) - Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[8].Value);
                }
            }
            else if (e.ColumnIndex == 10 || e.ColumnIndex == 4)
            {
                if (!dataGridView1.Rows[e.RowIndex].Cells[10].Value.ToString().Equals(string.Empty) && !dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString().Equals(string.Empty))
                {
                    dataGridView1.Rows[e.RowIndex].Cells[11].Value = Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString()) * Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[10].Value);
                }
            }
            else if (e.ColumnIndex == 10 || e.ColumnIndex == 8)
            {
                if (!dataGridView1.Rows[e.RowIndex].Cells[10].Value.ToString().Equals(string.Empty) && !dataGridView1.Rows[e.RowIndex].Cells[9].Value.ToString().Equals(string.Empty))
                {
                    dataGridView1.Rows[e.RowIndex].Cells[13].Value = (Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[10].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[8].Value) - Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[12].Value)) / Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[2].Value);
                }
            }
            if (e.ColumnIndex == 11)
            {
                Dictionary<string, string> dic = new Dictionary<string, string>();
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1.Rows[i].Cells[0].Value != null && !dataGridView1.Rows[i].Cells[11].Value.ToString().Equals(string.Empty))
                    {
                        string keys = "面料";
                        if (dic.ContainsKey(keys))
                        {
                            dic[keys] = (Convert.ToInt32(dic[keys].Split('=')[0]) + Convert.ToInt32(dataGridView1.Rows[i].Cells[11].Value.ToString())).ToString() + "=" + dic[keys].Split('=')[1];
                        }
                        else
                        {
                            dic.Add(keys, dataGridView1.Rows[i].Cells[11].Value.ToString() + "=" + i);
                        }
                    }
                }
                List<string> key = new List<string>(dic.Keys);
                double chushu = 0;
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1.Rows[i].Cells[2].Value != null)
                    {
                        chushu = chushu + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value.ToString());
                    }
                }
                if (chushu != 0)
                {
                    foreach (string k in key)
                    {
                        dataGridView1.Rows[Convert.ToInt32(dic[k].Split('=')[1])].Cells[14].Value = Convert.ToDouble(dic[k].Split('=')[0]) / chushu;
                    }
                }
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            List<HeSuan> ML = new List<HeSuan>();
            Boolean pd = false;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1[2, i].Value != null && dataGridView1[2, i].Value.ToString() != "")
                {
                    pd = true;
                    if (dataGridView1[0, i].Value != null)
                    {
                        HeSuan h = new HeSuan()
                        {
                            LOT = dataGridView1[0, i].Value.ToString(),
                            订单数量 = dataGridView1[1, i].Value.ToString(),
                            //Name = "面料", 
                            实际出口数量 = dataGridView1[2, i].Value.ToString(),
                            色号颜色 = dataGridView1[3, i].Value.ToString(),
                            单价 = dataGridView1[4, i].Value.ToString(),
                            预计单耗 = dataGridView1[5, i].Value.ToString(),
                            预计成本 = dataGridView1[6, i].Value.ToString(),
                            预计用量 = dataGridView1[7, i].Value.ToString(),
                            库存 = dataGridView1[8, i].Value.ToString(),
                            订量 = dataGridView1[9, i].Value.ToString(),
                            实际到货量 = dataGridView1[10, i].Value.ToString(),
                            实际到货金额 = dataGridView1[11, i].Value.ToString(),
                            剩余数量 = dataGridView1[12, i].Value.ToString(),
                            平均单耗 = dataGridView1[13, i].Value.ToString(),
                            结算成本 = dataGridView1[14, i].Value.ToString(),
                        };
                        ML.Add(h);
                    }
                }
            }
            if (pd) 
            {
                MessageBox.Show("实际出口数量为空 无法生存面辅料订购单");
            }else{
            InputCreatYjcb icb = new InputCreatYjcb(fm, cdhao, ML);
            icb.Show();}
        }
    }
}
