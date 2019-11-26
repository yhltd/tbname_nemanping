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

namespace PurchasingProcedures
{
    public partial class mflDgd : Form
    {
        protected clsAllnewLogic cal;
        protected GongNeng2 gn;
        //private string dh;
        protected Define1 df; 
        public List<HeSuan> hesuan;
        //public string pinming;
        private string cd;
        private string jiagongchang;
        private string kuanshi;
        public List<clsBuiness.MianFuLiaoDingGouDan> ChuanHuiMFL;
        public mflDgd(List<HeSuan> hs,string caidan,string jgc ,string ks)
        {
            hesuan = hs.GroupBy(g => new { g.Name,g.LOT}).Select(s=>s.First()).ToList<clsBuiness.HeSuan>();
            //dh = CdHao;
            cd = caidan;
            ChuanHuiMFL = new List<clsBuiness.MianFuLiaoDingGouDan>();
            cal = new clsAllnewLogic();
            gn = new GongNeng2();
            df = new Define1();
            jiagongchang = jgc;
            kuanshi = ks;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;

            InitializeComponent();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        public void mflDgd_Load(object sender, EventArgs e)
        {
            
            DataTable dt = new DataTable();
            for (int i = 0; i < dataGridView1.ColumnCount; i++) 
            {
                    dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString());
            }
            //dataGridView1.DataSource = null;
            dataGridView1.DataSource = dt;
            if ( ChuanHuiMFL.Count<=0)
            {
                List<MianFuLiaoDingGouDan> mfldgd = new List<MianFuLiaoDingGouDan>();
                foreach (HeSuan h in hesuan)
                {
                    List<clsBuiness.GongHuoFang> ghf = df.selectGongHuoFang().FindAll(gh => gh.HuoHao.Equals(h.LOT)).ToList<clsBuiness.GongHuoFang>();
                    clsBuiness.GongHuoFang gonghuofang = ghf.Find(g => g.SeHao.Equals(h.色号颜色.Split(' ')[0]));
                    if (gonghuofang != null)
                    {
                        MianFuLiaoDingGouDan mfl = new MianFuLiaoDingGouDan()
                        {
                            PingMing = h.Name,
                            HuoHao = h.LOT,
                            SeHao = h.色号颜色.Split(' ')[0],
                            YanSe = h.色号颜色.Split(' ')[1],
                            GuiGe = gonghuofang.Guige,
                            DanWei = "单位",
                            DanJia = gonghuofang.DanJia,
                            ShuLiang = h.订单数量,
                            ZongJinE = (Convert.ToDouble(gonghuofang.DanJia.ToString()) * Convert.ToInt32(h.订单数量.ToString())).ToString(),
                            CaiDanHao = cd,
                            GongFang = gonghuofang.GongHuoFangA + "-" + gonghuofang.GongHuoFangB,
                        };
                        mfldgd.Add(mfl);
                    }
                    else
                    {
                        MessageBox.Show("生成失败！ 原因:供货表里没有该 色号的信息");
                    }
                }
                foreach (MianFuLiaoDingGouDan mf in mfldgd)
                {
                    dt.Rows.Add(string.Empty, mf.PingMing, mf.HuoHao, mf.SeHao, mf.YanSe, mf.GuiGe, mf.DanWei, mf.DanJia, mf.ShuLiang, mf.ZongJinE, mf.CaiDanHao);
                }
                dataGridView1.DataSource = dt;
                txt_gongfang.Text = mfldgd[0].GongFang;
                //label7.Text = pinming;
            }
            else 
            {
                foreach (MianFuLiaoDingGouDan mf in ChuanHuiMFL)
                {
                    dt.Rows.Add(mf.Id, mf.PingMing, mf.HuoHao, mf.SeHao, mf.YanSe, mf.GuiGe, mf.DanWei, mf.DanJia, mf.ShuLiang, mf.ZongJinE, mf.CaiDanHao);
                }
                dataGridView1.DataSource = dt;
                if (ChuanHuiMFL != null && ChuanHuiMFL.Count > 0)
                {
                    txt_gongfang.Text = ChuanHuiMFL[0].GongFang;
                    txt_XuFang.Text = ChuanHuiMFL[0].XuFang;
                    txt_hetonghao.Text = ChuanHuiMFL[0].HeTongHao;
                    txt_shijian.Text = ChuanHuiMFL[0].QianYueShiJian;
                    txt_didian.Text = ChuanHuiMFL[0].QianYueDiDan;
                    //label7.Text = pinming;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try 
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("Id");
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("Id"))
                    {
                        dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                    }
                }
                dt.Columns.Add("供方");
                dt.Columns.Add("需方");
                dt.Columns.Add("合同号");
                dt.Columns.Add("签约时间");
                dt.Columns.Add("签约地点");
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1.Rows[i].Cells[6].Value != null)
                    {
                        dt.Rows.Add(dataGridView1.Rows[i].Cells[0].Value,dataGridView1.Rows[i].Cells[1].Value, dataGridView1.Rows[i].Cells[2].Value, dataGridView1.Rows[i].Cells[3].Value, dataGridView1.Rows[i].Cells[4].Value, dataGridView1.Rows[i].Cells[5].Value, dataGridView1.Rows[i].Cells[6].Value, dataGridView1.Rows[i].Cells[7].Value, dataGridView1.Rows[i].Cells[8].Value, dataGridView1.Rows[i].Cells[9].Value, dataGridView1.Rows[i].Cells[10].Value, txt_gongfang.Text, txt_XuFang.Text, txt_hetonghao.Text, txt_shijian.Text, txt_didian.Text);
                    }
                }
                cal.SaveMianFuliaoDingGouDan(dt);
                MessageBox.Show("保存成功！");
            }
            catch (Exception ex) 
            {
                throw ex; 

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            PingMingSelect pms = new PingMingSelect(this,cd);
            pms.Show();
            this.Hide();
             
        }

        //private void button1_Click(object sender, EventArgs e)
        //{

        //    List<clsBuiness.CaiDan> caidanlist = gn.selectCaiDan(dh);
        //    List<HeSuan> hs = new List<HeSuan>();
        //    Dictionary<string, string> dic = new Dictionary<string, string>();
        //    List<PeiSe> peiselist = cal.selectPeise("").FindAll(p => p.PingMing.Equals(comboBox1.Text));
        //    List<MianFuLiaoDingGouDan> mfldgd = new List<MianFuLiaoDingGouDan>();
        //    foreach (PeiSe ps in peiselist)
        //    {
        //        foreach (clsBuiness.CaiDan cd in caidanlist)
        //        {
        //            if (cd.ART.Equals(ps.HuoHao))
        //            {
        //                if (dic.ContainsKey(cd.LOT))
        //                {
        //                    dic[cd.LOT] = (Convert.ToInt32(dic[cd.LOT]) + Convert.ToInt32(cd.Sub_Total)).ToString();
        //                }
        //                else
        //                {
        //                    dic.Add(cd.LOT, cd.Sub_Total);
        //                }
        //            }
        //        }
        //    }
        //    //end 获取配色表里的信息 加以计算（色号+总数）
        //    foreach (clsBuiness.CaiDan c in caidanlist)
        //    {
        //        if (dic.ContainsKey(c.LOT))
        //        {
        //            if (dic[c.LOT].Split('=').Length < 2)
        //            {
        //                dic[c.LOT] = dic[c.LOT] + "=" + c.COLORID.Trim() + c.COLOR.Trim(); //色号&颜色
        //            }
        //            else
        //            {
        //                dic[c.LOT] = dic[c.LOT] + "=" + " ";
        //            }
        //        }
        //    }
        //    List<string> key = new List<string>(dic.Keys);
        //    List<MianFuLiaoDingGouDan> endlist = new List<MianFuLiaoDingGouDan>();
        //    for (int i = 0; i < key.Count; i++)
        //    {
        //        //MianFuLiaoDingGouDan mfl = new MianFuLiaoDingGouDan();
        //        //mfl. = key[i];
        //        //if (dic[key[i]].Split('=').Length > 0)
        //        //{
        //        //    mfl.订单数量 = dic[key[i]].Split('=')[0];
        //        //}
        //        //mfl.Name = key[i];
        //        //if (dic[key[i]].Split('=').Length > 1)
        //        //{
        //        //    mfl.色号颜色 = dic[key[i]].Split('=')[1];
        //        //}

        //    }

        //}
    }
}
