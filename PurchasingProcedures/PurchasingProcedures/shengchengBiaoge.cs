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
using System.IO;
using Spire.Xls;
namespace PurchasingProcedures
{

    public partial class shengchengBiaoge : Form
    {
        private string cdNo;

        protected GongNeng2 gn;
        protected clsAllnewLogic cal;
        private string foldPath;
        private int lie;
        private string STYLE;
        private List<string> color;
        private string insertStr;
        private List<HeSuan> mllist;
        private List<HeSuan> Fuliao;
        public shengchengBiaoge(string caidanNo, List<HeSuan> ml, List<HeSuan> fuliao)
        {
            InitializeComponent();
            cdNo = caidanNo;
            mllist = ml;
            Fuliao = fuliao.GroupBy(g => g.Name.Trim()).Select(s => s.First()).ToList<HeSuan>();
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            gn = new GongNeng2();
            cal = new clsAllnewLogic();
            List<clsBuiness.CaiDan> cdlist = gn.selectCaiDan("").GroupBy(c => c.LOT).Select(s => s.First()).ToList<clsBuiness.CaiDan>();
            STYLE = cdlist[0].STYLE;

        }

        private void shengchengBiaoge_Load(object sender, EventArgs e)
        {
            //配色
            CreatePeiSe();
            //单耗
            CreateDanHao();
            //核定
            CreateHeding();
        }

        private void CreateDanHao()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("面料", typeof(string));
            dt.Columns.Add("货号", typeof(string));
            dt.Columns.Add("幅宽", typeof(string));
            dt.Columns.Add("色号&颜色", typeof(string));
            dt.Columns.Add("单耗", typeof(string));
            dt.Columns.Add("单价", typeof(string));
            dt.Columns.Add("金额", typeof(string));
            List<DanHao> dhlist = cal.SelectDanHao("").Where(c => c.CaiDanNo.Trim().ToUpper().Equals(cdNo.Trim().ToUpper())).ToList<DanHao>();
            foreach (DanHao dh in dhlist)
            {
                dt.Rows.Add(dh.Name, dh.HuoHao, dh.GuiGe, dh.Yanse, dh.DanHao1, dh.Danjia, dh.Jine);
            }
            dgv_dh.DataSource = dt;
        }
        private void CreatePeiSe()
        {
            color = new List<string>();
            lie = 0;
            List<clsBuiness.CaiDan> cdlist = gn.selectCaiDan("").GroupBy(c => c.LOT).Select(s => s.First()).ToList<clsBuiness.CaiDan>();
            DataTable dt = new DataTable();
            dt.Columns.Add("品名", typeof(string));
            dt.Columns.Add("货号", typeof(string));
            dt.Columns.Add("规格/幅宽", typeof(string));
            string dgvSTR = "";
            foreach (clsBuiness.CaiDan cd in cdlist)
            {
                dt.Columns.Add(cd.LOT, typeof(string));
                dgvSTR = dgvSTR + "=" + cd.LOT;
            }
            //List<clsBuiness.CaiDan> listcd = cdlist.FindAll(c => c.CaiDanHao.Equals(cdNo));
            List<clsBuiness.CaiDanALL> cball = new List<clsBuiness.CaiDanALL>();
            List<clsBuiness.CaiDanALL> cdall = new List<clsBuiness.CaiDanALL>();
            List<clsBuiness.CaiDan> caidanlist = gn.selectCaiDan(cdNo);
            List<clsBuiness.CaiDan_RGL2> caidanRGL2 = gn.selectCaiDanRGL2(cdNo);
            List<clsBuiness.CaiDan_SLIM> caidanSLIM = gn.selectCaiDanSLIM(cdNo);
            List<clsBuiness.CaiDan_RGLJ> caidanRGLJ = gn.selectCaiDanRGLJ(cdNo);
            List<clsBuiness.CaiDan_D_PANT> caidanD_PANT = gn.selectCaiDanD_PANT(cdNo);
            List<clsBuiness.CaiDan_C_PANT> caidanC_PANT = gn.selectCaiDanC_PANT(cdNo);
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

            cdall = cball.GroupBy(c => c.LOT).Select(s => s.First()).ToList<clsBuiness.CaiDanALL>();
            string mlNo = "";
            if (cdall != null && cdall.Count > 0)
            {
                mlNo = cdall[0].MianLiao;
            }
            List<clsBuiness.PeiSe> ps = cal.selectPeise("").FindAll(p => p.Fabrics.Trim().ToUpper().Equals(mlNo));
            insertStr = "";
            foreach (clsBuiness.PeiSe p in ps)
            {
                insertStr = p.PingMing + "=" + p.HuoHao + "=" + p.GuiGe;
                lie = 0;
                if (dgvSTR.Contains("61601C1"))
                {
                    insertStr = insertStr + "=" + p.C61601C1;
                    color.Add("黑 BLACK");
                    lie++;
                }
                if (dgvSTR.Contains("61602C1"))
                {
                    insertStr = insertStr + "=" + p.C61602C1;
                    color.Add("碳灰CHARCOAL");
                    lie++;
                }
                if (dgvSTR.Contains("61603C1"))
                {
                    insertStr = insertStr + "=" + p.C61603C1;
                    color.Add("海军蓝 NAVY");
                    lie++;
                }
                if (dgvSTR.Contains("61605C1"))
                {
                    insertStr = insertStr + "=" + p.C61605C1;
                    color.Add("米色 TAN");
                    lie++;
                }
                if (dgvSTR.Contains("61606C1"))
                {
                    insertStr = insertStr + "=" + p.C61606C1;
                    color.Add("灰色 GREY");
                    lie++;
                }
                if (dgvSTR.Contains("61607C1"))
                {
                    insertStr = insertStr + "=" + p.C61607C1;
                    color.Add("银灰色 SILVER GREY");
                    lie++;
                }
                if (dgvSTR.Contains("61609C1"))
                {
                    insertStr = insertStr + "=" + p.C61609C1;
                    color.Add("棕色 BROWN");
                    lie++;
                }
                if (dgvSTR.Contains("61611C1"))
                {
                    insertStr = insertStr + "=" + p.C61601C1;
                    color.Add("枣红色 BURGUNDY");
                    lie++;
                }
                if (dgvSTR.Contains("61618C1"))
                {
                    insertStr = insertStr + "=" + p.C61618C1;
                    color.Add("蓝色 BLUE");
                    lie++;
                }
                if (dgvSTR.Contains("61624C1"))
                {
                    insertStr = insertStr + "=" + p.C61624C1;
                    color.Add("橄榄绿 OLIVE");
                    lie++;
                }
                if (dgvSTR.Contains("61627C1"))
                {
                    insertStr = insertStr + "=" + p.C61627C1;
                    color.Add("钴蓝色 COBALT BLUE");
                    lie++;
                }
                if (dgvSTR.Contains("61631C1"))
                {
                    insertStr = insertStr + "=" + p.C61631C1;
                    color.Add("紫罗兰 IMPERIAL PURPLE");
                    lie++;
                }
                if (dgvSTR.Contains("61632C1"))
                {
                    insertStr = insertStr + "=" + p.C61632C1;
                    color.Add("宝石蓝 ROYAL BLUE");
                    lie++;
                }
                if (dgvSTR.Contains("61633C1"))
                {
                    insertStr = insertStr + "=" + p.C61633C1;
                    color.Add("苋红 RUMBA RED");
                    lie++;
                }
                if (dgvSTR.Contains("61634C1"))
                {
                    insertStr = insertStr + "=" + p.C61634C1;
                    color.Add("鹿褐色 GOLDEN BROWN");
                    lie++;
                }

                dt.Rows.Add(insertStr.Split('='));
            }

            dgv_ps.DataSource = dt;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                ShengChengBiaoGeXuanZe scbgz = new ShengChengBiaoGeXuanZe("打印", dgv_ps, dgv_dh, dataGridView1, color, lie, STYLE, cdNo);
                scbgz.ShowDialog();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public void CreateHeding()
        {
            //List<clsBuiness.CaiDan> cdlist = gn.selectCaiDan(cdNo);
            DataTable dt = new DataTable();
            //dt.Columns.Add("面辅料结算成本", typeof(string));
            dt.Columns.Add("类型", typeof(string));
            //foreach (clsBuiness.HeSuan c in Fuliao)
            //{
            //    dt.Columns.Add(c.Name);
            //}
            dt.Columns.Add("服装数量", typeof(string));
            dt.Columns.Add("单价成本", typeof(string));
            //dt.Columns.Add("总计", typeof(string));


            dt.Rows.Add("面辅料结算成本", mllist.Sum(s =>Convert.ToInt32(s.订单数量)), mllist[0].结算成本);//
            foreach (HeSuan hs in Fuliao)
            {
                dt.Rows.Add(hs.Name, hs.订单数量, hs.结算成本);
            }
            //foreach
            dataGridView1.DataSource = dt;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("确认要保存‘单耗’‘配色’ ‘核定成本’三份表格吗？", "系统提示", MessageBoxButtons.YesNo);
            if (dr == DialogResult.Yes)
            {
                try
                {
                    FolderBrowserDialog dialog = new FolderBrowserDialog();
                    dialog.Description = "请选择文件路径";
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        foldPath = dialog.SelectedPath;
                        CreateExcel(foldPath);
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
                ShengChengBiaoGeXuanZe scbgz = new ShengChengBiaoGeXuanZe("保存", dgv_ps, dgv_dh, dataGridView1, color, lie, STYLE, cdNo);
                scbgz.ShowDialog();

            }
        }



        private void CreateExcel(string path)
        {
            DataTable dt = new DataTable();

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



            DataTable dt2 = new DataTable();

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

            DataTable dt3 = new DataTable();
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
            gn.SavePeiSeToExcel(dt, dt2, dt3, path, STYLE, cdNo);
            foldPath = path + "\\配色表-" + STYLE + "-" + cdNo + ".xls";
            //foldPath2 = path + "\\单耗-" + STYLE + "-" + cdNo + ".xls";
        }




    }
}
