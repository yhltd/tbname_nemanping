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
        private string dh;
        public mflDgd(string CdHao)
        {
            dh = CdHao;
            cal = new clsAllnewLogic();
            gn = new GongNeng2();
            InitializeComponent();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void mflDgd_Load(object sender, EventArgs e)
        {
            List<DanHao> list = cal.SelectDanHao("").FindAll(d => d.CaiDanNo.Trim().Equals(dh) && d.Type.Equals("辅料")).GroupBy(gp => gp.Name.Trim()).Select(s => s.First()).ToList<DanHao>();
            //foreach (DanHao dh in list)
            //{
            //    ToolStripMenuItem ms = new ToolStripMenuItem();
            //    ms.Name = dh.Type.ToString();
            //    ms.Text = dh.Name.ToString();
            //    ms_pingming.Items.Add(ms);
            //}
            comboBox1.DataSource = list;
            comboBox1.DisplayMember = "Name";
            comboBox1.ValueMember = "Id";

        }
        //private List<HeSuan> CreateFuLiao()
        //{
        //    //获取配色表里的信息 加以计算（色号+总数）start
            
        //    List<clsBuiness.GongHuoFang> ghflist = df.selectGongHuoFang();
        //    // foreach (KeyValuePair<string,string>kvp in dic)
        //    //{               

        //    //}
        //    List<clsBuiness.DanHao> dhlist = cal.SelectDanHao("");
        //    List<clsBuiness.KuCun> kclist = cal.SelectKC();
        //    List<string> key = new List<string>(dic.Keys);

        //    for (int i = 0; i < key.Count; i++)
        //    {
        //        foreach (clsBuiness.GongHuoFang g in ghflist)
        //        {
        //            if (dic.ContainsKey(key[i]))
        //            {
        //                if (dic[key[i]].Split('=')[1].Equals(g.SeHao.Trim() + g.Yanse.Trim()) && g.PingMing.Trim().Equals(e.ClickedItem.Text.Trim()))
        //                {
        //                    dic[key[i]] = dic[key[i]] + "=" + g.DanJia;//单价
        //                }
        //                else
        //                {
        //                    dic[key[i]] = dic[key[i]] + "=" + "";
        //                }

        //            }
        //        }
        //        foreach (clsBuiness.KuCun kc in kclist)
        //        {
        //            if (dic.ContainsKey(key[i]))
        //            {
        //                if (kc.PingMing.Trim().Equals(e.ClickedItem.Text.Trim()) && kc.SeHao.Trim().Equals(dic[key[i]].Split('=')[1]))
        //                {
        //                    dic[key[i]] = dic[key[i]] + "=" + kc.ShuLiang;
        //                    dic[key[i]] = dic[key[i]] + "=" + (Convert.ToInt32(dic[key[i]].Split('=')[5].ToString()) - Convert.ToInt32(kc.ShuLiang.ToString())); //订量
        //                    dic[key[i]] = dic[key[i]] + "=" + " ";
        //                    //dic[kvp.Key] = dic[kvp.Key] + "=" +
        //                }
        //                else
        //                {
        //                    dic[key[i]] = dic[key[i]] + "=" + "";
        //                }
        //            }
        //        }

        //    }
        //    List<HeSuan> endlist = new List<HeSuan>();
        //    for (int i = 0; i < key.Count; i++)
        //    {
        //        HeSuan endhs = new HeSuan();
        //        endhs.LOT = key[i];
        //        if (dic[key[i]].Split('=').Length > 0)
        //        {
        //            endhs.订单数量 = dic[key[i]].Split('=')[0];
        //        }
        //        endhs.Name = key[i];
        //        if (dic[key[i]].Split('=').Length > 1)
        //        {
        //            endhs.色号颜色 = dic[key[i]].Split('=')[1];
        //        }
        //        if (dic[key[i]].Split('=').Length > 2)
        //        {
        //            endhs.单价 = dic[key[i]].Split('=')[2];
        //        }
        //        if (dic[key[i]].Split('=').Length > 3)
        //        {
        //            endhs.预计单耗 = dic[key[i]].Split('=')[3];
        //        }
        //        if (dic[key[i]].Split('=').Length > 4)
        //        {
        //            endhs.预计成本 = dic[key[i]].Split('=')[4];
        //        }
        //        if (dic[key[i]].Split('=').Length > 5)
        //        {
        //            endhs.预计用量 = dic[key[i]].Split('=')[5];
        //        }
        //        if (dic[key[i]].Split('=').Length > 6)
        //        {
        //            endhs.库存 = dic[key[i]].Split('=')[6];
        //        }
        //        if (dic[key[i]].Split('=').Length > 7)
        //        {
        //            endhs.订量 = dic[key[i]].Split('=')[7];
        //        }
        //        if (dic[key[i]].Split('=').Length > 8)
        //        {
        //            endhs.实际到货量 = dic[key[i]].Split('=')[8];
        //        }

        //        if (endhs.实际到货量 != null && endhs.单价 != null && !endhs.实际到货量.Equals(string.Empty) && !endhs.单价.Equals(string.Empty))
        //        {
        //            endhs.实际到货金额 = (Convert.ToInt32(endhs.单价) * Convert.ToInt32(endhs.实际到货量)).ToString();
        //        }
        //        //if (!endhs.库存.Equals(string.Empty) && !endhs.实际到货量.Equals(string.Empty) && !endhs.剩余数量.Equals(string.Empty))
        //        //{
        //        //   endhs.平均单耗 = (endhs.单价-endhs.剩余数量+endhs.库存)
        //        //}
        //        endlist.Add(endhs);
        //    }
        //    return endlist;
        //}

        private void button1_Click(object sender, EventArgs e)
        {

            List<clsBuiness.CaiDan> caidanlist = gn.selectCaiDan(dh);
            List<HeSuan> hs = new List<HeSuan>();
            Dictionary<string, string> dic = new Dictionary<string, string>();
            List<PeiSe> peiselist = cal.selectPeise("").FindAll(p => p.PingMing.Equals(comboBox1.Text));
            List<MianFuLiaoDingGouDan> mfldgd = new List<MianFuLiaoDingGouDan>();
            foreach (PeiSe ps in peiselist)
            {
                foreach (clsBuiness.CaiDan cd in caidanlist)
                {
                    if (cd.ART.Equals(ps.HuoHao))
                    {
                        if (dic.ContainsKey(cd.LOT))
                        {
                            dic[cd.LOT] = (Convert.ToInt32(dic[cd.LOT]) + Convert.ToInt32(cd.Sub_Total)).ToString();
                        }
                        else
                        {
                            dic.Add(cd.LOT, cd.Sub_Total);
                        }
                    }
                }
            }
            //end 获取配色表里的信息 加以计算（色号+总数）
            foreach (clsBuiness.CaiDan c in caidanlist)
            {
                if (dic.ContainsKey(c.LOT))
                {
                    if (dic[c.LOT].Split('=').Length < 2)
                    {
                        dic[c.LOT] = dic[c.LOT] + "=" + c.COLORID.Trim() + c.COLOR.Trim(); //色号&颜色
                    }
                    else
                    {
                        dic[c.LOT] = dic[c.LOT] + "=" + " ";
                    }
                }
            }
            List<string> key = new List<string>(dic.Keys);
            List<MianFuLiaoDingGouDan> endlist = new List<MianFuLiaoDingGouDan>();
            for (int i = 0; i < key.Count; i++)
            {
                //MianFuLiaoDingGouDan mfl = new MianFuLiaoDingGouDan();
                //mfl. = key[i];
                //if (dic[key[i]].Split('=').Length > 0)
                //{
                //    mfl.订单数量 = dic[key[i]].Split('=')[0];
                //}
                //mfl.Name = key[i];
                //if (dic[key[i]].Split('=').Length > 1)
                //{
                //    mfl.色号颜色 = dic[key[i]].Split('=')[1];
                //}

            }

        }
    }
}
