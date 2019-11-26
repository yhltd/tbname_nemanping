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
            //try
            //{
                //



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

            foreach (HeSuan hs in list)
            {
                dt.Rows.Add(hs.Name,hs.LOT, hs.订单数量, "", hs.色号颜色, hs.单价, hs.预计单耗, hs.预计成本, hs.预计用量, hs.库存, hs.订量, hs.实际到货量, hs.实际到货金额, hs.剩余数量, hs.平均单耗, hs.结算成本);

            }
            dataGridView2.Columns.Clear();
            dataGridView2.DataSource = dt;
          



                
            //}
            //catch (Exception ex) { throw ex; }
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

        private List<HeSuan> CreateFuLiao(string name,string type)
        {
            List<HeSuan> endlist = new List<HeSuan>();
            List<clsBuiness.CaiDan> caidanlist = gn.selectCaiDan(cdhao);
            List<HeSuan> hs = new List<HeSuan>();
            List<DanHao> dh = cblist.FindAll(dc => dc.Type.Equals("辅料")); 
            Dictionary<string, string> dic = new Dictionary<string, string>();
            List<PeiSe> peis = cal.selectPeise("");
            List<clsBuiness.GongHuoFang> ghflist = df.selectGongHuoFang();
            // foreach (KeyValuePair<string,string>kvp in dic)
            //{               

            //}
            List<clsBuiness.DanHao> dhlist = cal.SelectDanHao("");
            List<clsBuiness.KuCun> kclist = cal.SelectKC();
            foreach(clsBuiness.DanHao dcc in dh){
                List<PeiSe> peiselist = peis.FindAll(p => p.PingMing.Equals(dcc.Name));
                //获取配色表里的信息 加以计算（色号+总数）start
                 List<string> key = new List<string>();
                 //dic.Keys = null;
                foreach (PeiSe ps in peiselist)
                {
                    foreach (clsBuiness.CaiDan cd in caidanlist)
                    {
                        if (cd.ART.Equals(ps.HuoHao))
                        {
                            if (dic.ContainsKey(dcc.Name + "=" + ps.HuoHao))
                            {
                                dic[dcc.Name+"="+ps.HuoHao] = (Convert.ToInt32(dic[dcc.Name+"="+ps.HuoHao]) + Convert.ToInt32(cd.Sub_Total)).ToString();
                            }
                            else
                            {
                                dic.Add(dcc.Name + "=" + ps.HuoHao, cd.Sub_Total);
                            }
                        }
                    }
                //end 获取配色表里的信息 加以计算（色号+总数）
                foreach (clsBuiness.CaiDan c in caidanlist)
                {
                    if (dic.ContainsKey(dcc.Name + "=" + ps.HuoHao))
                    {
                        if (dic[dcc.Name + "=" + ps.HuoHao].Split('=').Length < 2)
                        {
                            dic[dcc.Name + "=" + ps.HuoHao] = dic[dcc.Name + "=" + ps.HuoHao] + "=" + c.COLORID.Trim() + " " + c.COLOR.Trim(); //色号&颜色
                        }
                        else 
                        {
                            dic[dcc.Name + "=" + ps.HuoHao] = dic[dcc.Name + "=" + ps.HuoHao] + "=" + " ";
                        }
                    }
                }
                
                key = new List<string>(dic.Keys);

                for (int i = 0; i < key.Count; i++)
                {
                    foreach (clsBuiness.GongHuoFang g in ghflist)
                    {
                        if(dic.ContainsKey(key[i])){
                            if (dic[key[i]].Split('=')[1].Equals(g.SeHao.Trim() + g.Yanse.Trim()) && g.PingMing.Trim().Equals(name)) 
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
                hEsuan = endlist;

            
                }
            }
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

        private static DataTable DataGirdViewHeader(string name,string type)
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

        private List<HeSuan> GetHeSuanList(string name, string type)
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
                        if (dic[key[i]].Split('=')[1].Equals(g.SeHao.Trim() + g.Yanse.Trim()) && g.PingMing.Trim().Equals(name)) 
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
                        if (d.Style.Equals(ks) && d.JiaGongChang.Equals(jgc) && d.Name.Trim().Equals(name) && d.Yanse.Trim().Equals(dic[key[i]].Split('=')[1].Trim() + " " + dic[key[i]].Split('=')[2].Trim()))
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
                        if (kc.PingMing.Trim().Equals(name) && kc.SeHao.Trim().Equals(dic[key[i]].Split('=')[1]))
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
            if (hEsuan != null)
            {
                mflDgd mfl = new mflDgd(hEsuan, cdhao,jgc,ks);
                mfl.Show();

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
                Dictionary<string,string> dic = new Dictionary<string,string>();
                for (int i = 0; i < dataGridView2.Rows.Count; i++) 
                {
                    if (dataGridView2.Rows[i].Cells[12].Value!=null && !dataGridView2.Rows[i].Cells[12].Value.ToString().Equals(string.Empty))
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
                double chushu =0;
                for (int i = 0; i < dataGridView1.Rows.Count; i++) 
                {
                    if (dataGridView1.Rows[i].Cells[2].Value != null)
                    {
                        chushu = chushu + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value.ToString());
                    }
                }
                if(chushu != 0)
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
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1[0, i].Value != null)
                {
                    HeSuan h = new HeSuan()
                    {
                        LOT = dataGridView1[0, i].Value.ToString(),
                        订单数量 = dataGridView1[1, i].Value.ToString(),
                        //Name = "面料", 
                        实际出口数量 = dataGridView1[2,i].Value.ToString(),
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
            InputCreatYjcb icb = new InputCreatYjcb(this, cdhao, ML);
            icb.Show();
        }
    }
}
