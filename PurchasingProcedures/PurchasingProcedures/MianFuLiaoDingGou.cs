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

        public MianFuLiaoDingGou(string ml, string ks, string jgc, string cdh)
        {
            InitializeComponent();
            cal = new clsAllnewLogic();
            gn = new GongNeng2();
            df = new Define1();
            this.ml = ml;
            this.ks = ks;
            this.jgc = jgc;
            cdhao = cdh;
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void MianFuLiaoDingGou_Load(object sender, EventArgs e)
        {
            try 
            {
                List<DanHao> list = cal.SelectDanHao("").FindAll(d => d.CaiDanNo.Trim().Equals(cdhao)).GroupBy(gp => gp.Name.Trim()).Select(s => s.First()).ToList<DanHao>();
               foreach (DanHao dh in list) 
               {
                   ToolStripMenuItem ms = new ToolStripMenuItem();
                   ms.Name = dh.Type.ToString();
                   ms.Text = dh.Name.ToString() ;
                   this.ms_caidan.Items.Add(ms);
               }
            }
            catch (Exception ex) { throw ex; }
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
            try
            {
                string name = e.ClickedItem.Name;
                DataTable dt = DataGirdViewHeader(e);
                List<HeSuan> list = new List<HeSuan>();
                this.backgroundWorker1.RunWorkerAsync();
                JingDu form = new JingDu(this.backgroundWorker1, "计算中");// 显示进度条窗体
                form.ShowDialog(this);
                form.Close();
                if (name.Equals("辅料"))
                {
                   list= CreateFuLiao(e);
                }
                else
                {
                    list = GetHeSuanList(e);
                }
                
                foreach (HeSuan hs in list) 
                {
                    dt.Rows.Add(hs.LOT, hs.订单数量, "", hs.色号颜色, hs.单价, hs.预计单耗, hs.预计成本, hs.预计用量, hs.库存, hs.订量, hs.实际到货量, hs.实际到货金额, hs.剩余数量, hs.平均单耗, hs.结算成本);

                }
                dataGridView1.Columns.Clear();
                dataGridView1.DataSource = dt;
            }
            catch (Exception ex) 
            {
                throw ex;
            }
           

       
        }

        private List<HeSuan> CreateFuLiao(ToolStripItemClickedEventArgs e)
        {
            List<clsBuiness.CaiDan> caidanlist = gn.selectCaiDan(cdhao);
            List<HeSuan> hs = new List<HeSuan>();
            Dictionary<string, string> dic = new Dictionary<string, string>();
            List<PeiSe> peiselist = cal.selectPeise("").FindAll(p => p.PingMing.Equals(e.ClickedItem.Text));
            //获取配色表里的信息 加以计算（色号+总数）start
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
                    if(dic.ContainsKey(key[i])){
                        if (dic[key[i]].Split('=')[1].Equals(g.SeHao.Trim() + g.Yanse.Trim()) && g.PingMing.Trim().Equals(e.ClickedItem.Text.Trim())) 
                        {
                            dic[key[i]] = dic[key[i]] + "=" + g.DanJia;//单价
                        }else
                        {
                            dic[key[i]] = dic[key[i]] + "=" + "";
                        }
                    
                    }
                }
                foreach (clsBuiness.DanHao d in dhlist)
                {
                    if (dic.ContainsKey(key[i]))
                    {
                        if (d.Style.Equals(ks) && d.JiaGongChang.Equals(jgc) && d.Name.Trim().Equals(e.ClickedItem.Text.Trim()) && d.Yanse.Trim().Equals(dic[key[i]].Split('=')[1].Trim() + " " + dic[key[i]].Split('=')[2].Trim()))
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
                        if (kc.PingMing.Trim().Equals(e.ClickedItem.Text.Trim()) && kc.SeHao.Trim().Equals(dic[key[i]].Split('=')[1]))
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
            List<HeSuan> endlist = new List<HeSuan>();
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
                endlist.Add(endhs);
            }
            return endlist;
        }

        private static DataTable DataGirdViewHeader(ToolStripItemClickedEventArgs e)
        {
            string name = e.ClickedItem.Name;

            DataTable dt = new DataTable();
            dt.Columns.Add("LOT#", typeof(string));
            dt.Columns.Add("订单数量", typeof(string));
            dt.Columns.Add("实际出口数量", typeof(string));
            dt.Columns.Add("色号&颜色", typeof(string));
            dt.Columns.Add("单价", typeof(string));
            dt.Columns.Add(e.ClickedItem.Text + "预计单耗", typeof(string));
            dt.Columns.Add(e.ClickedItem.Text + "预计成本", typeof(string));
            dt.Columns.Add(e.ClickedItem.Text + "预计用量", typeof(string));
            dt.Columns.Add(e.ClickedItem.Text + "库存", typeof(string));
            dt.Columns.Add(e.ClickedItem.Text + "订量", typeof(string));
            dt.Columns.Add(e.ClickedItem.Text + "实际到货量", typeof(string));
            dt.Columns.Add(e.ClickedItem.Text + "实际到货金额", typeof(string));
            dt.Columns.Add(e.ClickedItem.Text + "剩余数量", typeof(string));
            dt.Columns.Add(e.ClickedItem.Text + "平均单耗", typeof(string));
            dt.Columns.Add(e.ClickedItem.Text + "结算成本", typeof(string));
            if (name.Equals("辅料"))
            {
                dt.Columns.Add("小计", typeof(string));
            }
            else
            {
                dt.Columns.Add("总数", typeof(string));
            }
            return dt;
        }

        private List<HeSuan> GetHeSuanList(ToolStripItemClickedEventArgs e)
        {
            List<clsBuiness.CaiDan> caidanlist = gn.selectCaiDan(cdhao);
            List<HeSuan> hs = new List<HeSuan>();
            Dictionary<string, string> dic = new Dictionary<string, string>();
            foreach (clsBuiness.CaiDan c in caidanlist)//成衣数量
            {
                if (dic.ContainsKey(c.LOT))
                {
                    dic[c.LOT] = (Convert.ToInt32(dic[c.LOT]) + Convert.ToInt32(c.Sub_Total)).ToString();
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
                        if (dic[key[i]].Split('=')[1].Equals(g.SeHao.Trim() + g.Yanse.Trim()) && g.PingMing.Trim().Equals(e.ClickedItem.Text.Trim())) 
                        {
                            dic[key[i]] = dic[key[i]] + "=" + g.DanJia;//单价
                        }else
                        {
                            dic[key[i]] = dic[key[i]] + "=" + "";
                        }
                    }
                }
                foreach (clsBuiness.DanHao d in dhlist)
                {
                    if(dic.ContainsKey(key[i]))
                    {
                        if (d.Style.Equals(ks) && d.JiaGongChang.Equals(jgc) && d.Name.Trim().Equals(e.ClickedItem.Text.Trim()) && d.Yanse.Trim().Equals(dic[key[i]].Split('=')[1].Trim() + " " + dic[key[i]].Split('=')[2].Trim()))
                        {
                            dic[key[i]] = dic[key[i]] + "=" + d.DanHao1;//预计单耗
                            dic[key[i]] = dic[key[i]] + "=" + Convert.ToInt32(dic[key[i]].Split('=')[2]) * Convert.ToInt32(d.DanHao1);//预计成本
                            dic[key[i]] = dic[key[i]] + "=" + Convert.ToInt32(dic[key[i]].Split('=')[0]) * Convert.ToInt32(d.DanHao1);//预计用量
                        }else
                        {
                            dic[key[i]] = dic[key[i]] + "=" + "";
                        }
                    }
                }
                foreach (clsBuiness.KuCun kc in kclist)
                {
                    if(dic.ContainsKey(key[i]))
                    {
                        if (kc.PingMing.Trim().Equals(e.ClickedItem.Text.Trim()) && kc.SeHao.Trim().Equals(dic[key[i]].Split('=')[1]))
                        {
                            dic[key[i]] = dic[key[i]] + "=" + kc.ShuLiang;
                            dic[key[i]] = dic[key[i]] + "=" + (Convert.ToInt32(dic[key[i]].Split('=')[5].ToString()) - Convert.ToInt32(kc.ShuLiang.ToString())); //订量
                            dic[key[i]] = dic[key[i]] + "=" + " ";
                            //dic[kvp.Key] = dic[kvp.Key] + "=" +
                        }else
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
            return list;
        }
    }
}
