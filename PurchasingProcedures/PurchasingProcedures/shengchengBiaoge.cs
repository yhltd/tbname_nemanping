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
        List<clsBuiness.CaiDanALL> cball;
        private int rowindex;
        private int cloumnindex;


        double sum实际出口数量 = 0;

        public shengchengBiaoge(string caidanNo, List<HeSuan> ml, List<HeSuan> fuliao, List<clsBuiness.CaiDanALL> cball1)
        {
            InitializeComponent();
            cball = new List<CaiDanALL>();
            cball = cball1;

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

            //核算表

            create_hesuanbiao();

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

        private void create_hesuanbiao()
        {
            try
            {
                // cball
                var qtyTable = new DataTable();
                int jk = 0;
                int cloumnindex = 1;

                #region 面料
                List<string> quchongnashuidanwei = (from v in mllist select v.LOT).Distinct().ToList();


                qtyTable.Columns.Add("LOT", System.Type.GetType("System.String"));//0
                for (int i = 0; i < quchongnashuidanwei.Count; i++)
                {
                    qtyTable.Columns.Add(quchongnashuidanwei[i], System.Type.GetType("System.String"));//0

                }
                qtyTable.Columns.Add("总数", System.Type.GetType("System.String"));//0


                qtyTable.Rows.Add(qtyTable.NewRow());
                qtyTable.Rows.Add(qtyTable.NewRow());
                qtyTable.Rows[jk][0] = "订单数量";
                qtyTable.Rows[jk + 1][0] = "实际出口数量";

                double zongshu = 0;
                for (int i = 0; i < quchongnashuidanwei.Count; i++)
                {

                    //
                    List<HeSuan> stockstate = mllist.FindAll(o => o.LOT == quchongnashuidanwei[i]);
                    List<string> quchongnashuidanwei1 = (from v in stockstate select v.色号颜色).Distinct().ToList();

                    for (int ii = 0; ii < quchongnashuidanwei1.Count; ii++)
                    {
                        HeSuan stockstate1 = stockstate.Find(o => o.色号颜色 == quchongnashuidanwei1[ii]);
                        List<clsBuiness.CaiDanALL> stockstate2 = cball.FindAll(o => o.LOT == quchongnashuidanwei[i] && quchongnashuidanwei1[ii].Contains(o.COLORID));
                        double allSub_Total = 0;
                        foreach (clsBuiness.CaiDanALL iyen in stockstate2)
                            if (iyen.Sub_Total.Length > 0)
                                allSub_Total = Convert.ToDouble(iyen.Sub_Total) + allSub_Total;

                        zongshu = zongshu + allSub_Total;
                        qtyTable.Rows[jk][cloumnindex] = allSub_Total;//订单数量

                        cloumnindex++;
                    }
                }
                //总数
                qtyTable.Rows[jk][cloumnindex] = zongshu;


                for (int j = 0; j < 200; j++)
                    qtyTable.Rows.Add(qtyTable.NewRow());


                jk = jk + 2;
                cloumnindex = 1;
                //面料
                string cloumn1 = "面料";
                //第一列
                hesuanbiao_cloumn1(qtyTable, jk, cloumn1);

                double sum预计成本 = 0;
                double sum预计用量 = 0;
                double sum库存 = 0;
                double sum订量 = 0;
                double sum实际到货量 = 0;
                double sum实际到货金额 = 0;
                double sum剩余数量 = 0;
                double sum平均单耗 = 0;
                double sum结算成本 = 0;


                for (int i = 0; i < quchongnashuidanwei.Count; i++)
                {

                    List<HeSuan> stockstate = mllist.FindAll(o => o.LOT == quchongnashuidanwei[i]);
                    List<string> quchongnashuidanwei1 = (from v in stockstate select v.色号颜色).Distinct().ToList();

                    for (int ii = 0; ii < quchongnashuidanwei1.Count; ii++)
                    {
                        HeSuan stockstate1 = mllist.Find(o => o.色号颜色 == quchongnashuidanwei1[ii] && o.LOT == quchongnashuidanwei[i]);


                        qtyTable.Rows[jk][cloumnindex] = quchongnashuidanwei[i];//面料
                        qtyTable.Rows[jk + 1][cloumnindex] = quchongnashuidanwei1[ii];//色号&颜色
                        qtyTable.Rows[jk + 2][cloumnindex] = stockstate1.单价;//单价
                        qtyTable.Rows[jk + 3][cloumnindex] = stockstate1.预计单耗;//单价
                        qtyTable.Rows[jk + 4][cloumnindex] = stockstate1.预计成本;//单价
                        qtyTable.Rows[jk + 5][cloumnindex] = stockstate1.预计用量;//单价
                        qtyTable.Rows[jk + 6][cloumnindex] = stockstate1.库存;//单价
                        qtyTable.Rows[jk + 7][cloumnindex] = stockstate1.订量;//单价
                        qtyTable.Rows[jk + 8][cloumnindex] = stockstate1.实际到货量;//单价
                        qtyTable.Rows[jk + 9][cloumnindex] = stockstate1.实际到货金额;//单价
                        qtyTable.Rows[jk + 10][cloumnindex] = stockstate1.剩余数量;//单价
                        qtyTable.Rows[jk + 11][cloumnindex] = stockstate1.平均单耗;//单价
                        qtyTable.Rows[jk + 12][cloumnindex] = stockstate1.结算成本;//单价


                        if (stockstate1.预计成本.Length > 0)
                            sum预计成本 = Convert.ToDouble(stockstate1.预计成本) + sum预计成本;
                        if (stockstate1.预计用量.Length > 0)
                            sum预计用量 = Convert.ToDouble(stockstate1.预计用量) + sum预计用量;
                        if (stockstate1.库存.Length > 0)
                            sum库存 = Convert.ToDouble(stockstate1.库存) + sum库存;
                        if (stockstate1.订量.Length > 0)
                            sum订量 = Convert.ToDouble(stockstate1.订量) + sum订量;
                        if (stockstate1.实际到货量.Length > 0 && stockstate1.实际到货量 != " ")
                            sum实际到货量 = Convert.ToDouble(stockstate1.实际到货量) + sum实际到货量;
                        if (stockstate1.实际到货金额.Length > 0)
                            sum实际到货金额 = Convert.ToDouble(stockstate1.实际到货金额) + sum实际到货金额;
                        if (stockstate1.剩余数量.Length > 0)
                            sum剩余数量 = Convert.ToDouble(stockstate1.剩余数量) + sum剩余数量;
                        if (stockstate1.平均单耗.Length > 0)
                            sum平均单耗 = Convert.ToDouble(stockstate1.平均单耗) + sum平均单耗;
                        if (stockstate1.结算成本.Length > 0)
                            sum结算成本 = Convert.ToDouble(stockstate1.结算成本) + sum结算成本;

                        cloumnindex++;
                    }
                    qtyTable.Rows[jk + 4][cloumnindex] = sum预计成本 / 3;
                    qtyTable.Rows[jk + 5][cloumnindex] = sum预计用量;//单价
                    qtyTable.Rows[jk + 6][cloumnindex] = sum库存;//单价
                    qtyTable.Rows[jk + 7][cloumnindex] = sum订量;//单价
                    qtyTable.Rows[jk + 8][cloumnindex] = sum实际到货量;//单价
                    qtyTable.Rows[jk + 10][cloumnindex] = sum剩余数量;//单价
                    if (sum实际出口数量 != 0)
                    {
                        qtyTable.Rows[jk + 11][cloumnindex] = Convert.ToString((sum库存 + sum实际到货量 - sum剩余数量) / sum实际出口数量);//单价

                        qtyTable.Rows[jk + 12][cloumnindex] = sum实际到货金额 / sum实际出口数量;//单价
                    }

                }
                jk = jk + 14;
                #endregion
                cloumnindex = 0;


                //辅料

                List<string> fuliao_lotlist = (from v in Fuliao select v.Name).Distinct().ToList();
                for (int i = 0; i < fuliao_lotlist.Count; i++)
                {
                    cloumnindex = 0;

                    cloumn1 = fuliao_lotlist[i];
                    //第一列
                    hesuanbiao_cloumn1(qtyTable, jk, cloumn1);

                    cloumnindex++;

                    List<HeSuan> stockstate = Fuliao.FindAll(o => o.Name == fuliao_lotlist[i]);//查找辅料名称

                    List<string> fuliao_lotlist1 = (from v in stockstate select v.LOT).Distinct().ToList();//此lot即配色表的货号

                    for (int i1 = 0; i1 < fuliao_lotlist1.Count; i1++)
                    {
                        List<HeSuan> stockstat1e = Fuliao.FindAll(o => o.LOT == fuliao_lotlist1[i1]);//查找辅料名称

                        List<string> quchongnashuidanwei1 = (from v in stockstat1e select v.色号颜色).Distinct().ToList();//根据这个货号下的颜色去重

                        for (int ii = 0; ii < quchongnashuidanwei1.Count; ii++)
                        {
                            HeSuan stockstate1 = Fuliao.Find(o => o.色号颜色 == quchongnashuidanwei1[ii] && o.LOT == fuliao_lotlist1[i1]);

                            qtyTable.Rows[jk][cloumnindex] = fuliao_lotlist1[i1];//辅料名称
                            qtyTable.Rows[jk + 1][cloumnindex] = quchongnashuidanwei1[ii];//色号&颜色
                            qtyTable.Rows[jk + 2][cloumnindex] = stockstate1.单价;//单价
                            qtyTable.Rows[jk + 3][cloumnindex] = stockstate1.预计单耗;//单价
                            qtyTable.Rows[jk + 4][cloumnindex] = stockstate1.预计成本;//单价
                            qtyTable.Rows[jk + 5][cloumnindex] = stockstate1.预计用量;//单价
                            qtyTable.Rows[jk + 6][cloumnindex] = stockstate1.库存;//单价
                            qtyTable.Rows[jk + 7][cloumnindex] = stockstate1.订量;//单价
                            qtyTable.Rows[jk + 8][cloumnindex] = stockstate1.实际到货量;//单价
                            qtyTable.Rows[jk + 9][cloumnindex] = stockstate1.实际到货金额;//单价
                            qtyTable.Rows[jk + 10][cloumnindex] = stockstate1.剩余数量;//单价
                            qtyTable.Rows[jk + 11][cloumnindex] = stockstate1.平均单耗;//单价
                            qtyTable.Rows[jk + 12][cloumnindex] = stockstate1.结算成本;//单价



                            if (stockstate1.预计成本.Length > 0)
                                sum预计成本 = Convert.ToDouble(stockstate1.预计成本) + sum预计成本;
                            if (stockstate1.预计用量.Length > 0)
                                sum预计用量 = Convert.ToDouble(stockstate1.预计用量) + sum预计用量;
                            if (stockstate1.库存.Length > 0)
                                sum库存 = Convert.ToDouble(stockstate1.库存) + sum库存;
                            if (stockstate1.订量.Length > 0)
                                sum订量 = Convert.ToDouble(stockstate1.订量) + sum订量;
                            if (stockstate1.实际到货量.Length > 0 && stockstate1.实际到货量 != " ")
                                sum实际到货量 = Convert.ToDouble(stockstate1.实际到货量) + sum实际到货量;
                            if (stockstate1.实际到货金额.Length > 0)
                                sum实际到货金额 = Convert.ToDouble(stockstate1.实际到货金额) + sum实际到货金额;
                            if (stockstate1.剩余数量.Length > 0)
                                sum剩余数量 = Convert.ToDouble(stockstate1.剩余数量) + sum剩余数量;
                            if (stockstate1.平均单耗.Length > 0)
                                sum平均单耗 = Convert.ToDouble(stockstate1.平均单耗) + sum平均单耗;
                            if (stockstate1.结算成本.Length > 0)
                                sum结算成本 = Convert.ToDouble(stockstate1.结算成本) + sum结算成本;

                            cloumnindex++;

                        }
                        qtyTable.Rows[jk][cloumnindex] = "小计";
                        qtyTable.Rows[jk + 4][cloumnindex] = sum预计成本 / 3;
                        qtyTable.Rows[jk + 5][cloumnindex] = sum预计用量;//单价
                        qtyTable.Rows[jk + 6][cloumnindex] = sum库存;//单价
                        qtyTable.Rows[jk + 7][cloumnindex] = sum订量;//单价
                        qtyTable.Rows[jk + 8][cloumnindex] = sum实际到货量;//单价
                        qtyTable.Rows[jk + 10][cloumnindex] = sum剩余数量;//单价
                        if (sum实际出口数量 != 0)
                        {
                            qtyTable.Rows[jk + 11][cloumnindex] = Convert.ToString((sum库存 + sum实际到货量 - sum剩余数量) / sum实际出口数量);//单价

                            qtyTable.Rows[jk + 12][cloumnindex] = sum实际到货金额 / sum实际出口数量;//单价
                        }

                    }

                    jk = jk + 14;
                }

                this.bindingSource1.DataSource = qtyTable;
                this.dataGridView2.DataSource = this.bindingSource1;

            }
            catch (Exception ex)
            {
                MessageBox.Show("数据异常：" + ex);

                throw ex;
            }

        }

        private static void hesuanbiao_cloumn1(DataTable qtyTable, int jk, string cloumn1)
        {
            qtyTable.Rows[jk][0] = cloumn1;
            qtyTable.Rows[jk + 1][0] = "色号&颜色";
            qtyTable.Rows[jk + 2][0] = "单价";
            qtyTable.Rows[jk + 3][0] = cloumn1 + "预计单耗";
            qtyTable.Rows[jk + 4][0] = cloumn1 + "预计成本";
            qtyTable.Rows[jk + 5][0] = cloumn1 + "预计用量";
            qtyTable.Rows[jk + 6][0] = cloumn1 + "库存";
            qtyTable.Rows[jk + 7][0] = cloumn1 + "订量";
            qtyTable.Rows[jk + 8][0] = cloumn1 + "实际到货量";
            qtyTable.Rows[jk + 9][0] = cloumn1 + "实际到货金额";
            qtyTable.Rows[jk + 10][0] = cloumn1 + "剩余数量";
            qtyTable.Rows[jk + 11][0] = cloumn1 + "平均单耗";
            qtyTable.Rows[jk + 12][0] = cloumn1 + "结算成本";
        }


        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                ShengChengBiaoGeXuanZe scbgz = new ShengChengBiaoGeXuanZe("打印", dgv_ps, dgv_dh, dataGridView1, color, lie, STYLE, cdNo, dataGridView2);
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


            dt.Rows.Add("面辅料结算成本", mllist.Sum(s => Convert.ToInt32(s.订单数量)), mllist[0].结算成本);//
            foreach (HeSuan hs in Fuliao)
            {
                dt.Rows.Add(hs.Name, hs.订单数量, hs.结算成本);
            }
            //foreach
            dataGridView1.DataSource = dt;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("确认要保存‘单耗’‘配色’ ‘核定成本’‘核算单’三份表格吗？", "系统提示", MessageBoxButtons.YesNo);
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
                ShengChengBiaoGeXuanZe scbgz = new ShengChengBiaoGeXuanZe("保存", dgv_ps, dgv_dh, dataGridView1, color, lie, STYLE, cdNo, dataGridView2);
                scbgz.ShowDialog();

            }
        }



        private void CreateExcel(string path)
        {
            try
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
                            str = str + "=" + dgv_ps.Rows[i].Cells[j].Value;//品名=货号=规格/幅宽=61605C1=61607C1=61609C1=61601C1=61627C1=61634C1
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
                            str = dataGridView1.Rows[i].Cells[0].Value + "=" + dataGridView1.Rows[i].Cells[1].Value + "=" + dataGridView1.Rows[i].Cells[2].Value + "=" + (Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value));
                        }
                        else
                        {
                            str = dataGridView1.Rows[i].Cells[0].Value + "=" + dataGridView1.Rows[i].Cells[1].Value + "=" + dataGridView1.Rows[i].Cells[2].Value + "=" + 0;
                        }
                        dt3.Rows.Add(str.Split('='));
                    }
                }

                //核算表

                DataTable dt4 = new DataTable();
                for (int i = 0; i < dataGridView2.Columns.Count; i++)
                {
                    if (!dataGridView2.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                    {
                        dt4.Columns.Add(dataGridView2.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                    }
                }
                string row1 = "";
                for (int i = 0; i < dataGridView2.ColumnCount; i++)
                {
                    if (row1 != "")
                        row1 = row1 + "=" + dataGridView2.Columns[i].HeaderCell.Value.ToString();
                    else
                        row1 = dataGridView2.Columns[i].HeaderCell.Value.ToString();

                }
                dt4.Rows.Add(row1.Split('='));

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
                //old4
                //gn.SavePeiSeToExcel(dt, dt2, dt3, path, STYLE, cdNo);
                //new
                gn.SavePeiSeToExcel(dt, dt2, dt3, path, STYLE, cdNo, dt4);
                foldPath = path + "\\配色表-" + STYLE + "-" + cdNo + ".xls";
                //foldPath2 = path + "\\单耗-" + STYLE + "-" + cdNo + ".xls";
            }
            catch (Exception ex)
            {
                MessageBox.Show("数据异常：" + ex);

                throw;
            }
        }

        private void dataGridView2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                double sjcksl_实际出口数量 = 0;
                int lastcloumn = dataGridView2.Columns.Count - 1;

                if (dataGridView2.Rows[2].Cells[dataGridView2.Columns.Count - 1].Value != null && dataGridView2.Rows[1].Cells["总数"].EditedFormattedValue != null && dataGridView2.Rows[1].Cells["总数"].EditedFormattedValue.ToString().Length>0)
                    sjcksl_实际出口数量 = Convert.ToDouble(dataGridView2.Rows[1].Cells["总数"].EditedFormattedValue.ToString());


                double kucun = 0;
                double shijidaohuoliang = 0;
                double shijidaohuojine = 0;
                double shengyushuliang = 0;
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {

                    if (dataGridView2.Rows[i].Cells[0].Value != null)
                    {
                        string ax = dataGridView2.Rows[i].Cells[0].Value.ToString();

                   
                        if (ax.Contains("库存"))
                        {
                            if (dataGridView2.Rows[i].Cells[lastcloumn].Value != null && dataGridView2.Rows[i].Cells[lastcloumn].Value.ToString().Length > 0)
                                kucun = Convert.ToDouble(dataGridView2.Rows[i].Cells[lastcloumn].Value.ToString());

                        }
                        else if (ax.Contains("实际到货量"))
                        {
                            if (dataGridView2.Rows[i].Cells[lastcloumn].Value != null && dataGridView2.Rows[i].Cells[lastcloumn].Value.ToString().Length > 0)
                                shijidaohuoliang = Convert.ToDouble(dataGridView2.Rows[i].Cells[lastcloumn].Value.ToString());

                        }
                        else if (ax.Contains("实际到货金额"))
                        {
                            if (dataGridView2.Rows[i].Cells[lastcloumn].Value != null && dataGridView2.Rows[i].Cells[lastcloumn].Value.ToString().Length>0)
                                shijidaohuojine = Convert.ToDouble(dataGridView2.Rows[i].Cells[lastcloumn].Value.ToString());

                        }
                        else if (ax.Contains("剩余数量"))
                        {
                            if (dataGridView2.Rows[i].Cells[lastcloumn].Value != null && dataGridView2.Rows[i].Cells[lastcloumn].Value.ToString().Length > 0)
                                shengyushuliang = Convert.ToDouble(dataGridView2.Rows[i].Cells[lastcloumn].Value.ToString());

                        }
                        else if (ax.Contains("平均单耗"))
                        {
                          //  if (dataGridView2.Rows[i].Cells[lastcloumn].Value != null && dataGridView2.Rows[i].Cells[lastcloumn].Value.ToString().Length > 0)
                            dataGridView2.Rows[i].Cells[lastcloumn].Value =   String.Format("{0:N2}",(Convert.ToDouble(kucun + shijidaohuoliang - shengyushuliang) / sjcksl_实际出口数量));
                           
                        }
                        else if (ax.Contains("结算成本"))
                        {

                            dataGridView2.Rows[i].Cells[lastcloumn].Value = String.Format("{0:N2}", Convert.ToDouble(shijidaohuojine) / sjcksl_实际出口数量);
                            kucun = 0;
                            shijidaohuoliang = 0;
                            shijidaohuojine = 0;
                            shengyushuliang = 0;
                        }
                    }

                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("数据异常：" + ex);
                return;


                throw;
            }


        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            rowindex = e.RowIndex;
            cloumnindex = e.ColumnIndex;


        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int i= this.tabControl1.SelectedIndex;
            if (i == 3)
            {

                label1.Text = "注意：" + "此界面内容需求部分按照实际手动填写(1.结算成本=实际到货金额/实际出口数量 2.平均单耗=(库存+实际到货量-剩余数量)/实际出口数量)";
 
            }
            else
                label1.Text = "";


        }




    }
}
