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
using System.Threading;
namespace PurchasingProcedures
{
    public partial class GongHuoFang : Form
    {
        protected DataTable dt;
        protected List<clsBuiness.GongHuoFang> list2;
        protected Define1 cal1;
        public GongHuoFang()
        {
            InitializeComponent();
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            cal1 = new Define1();
            list2 = new List<clsBuiness.GongHuoFang>();
        }

        #region 刷新按钮
        private void toolStripLabel5_Click(object sender, EventArgs e)
        {
            try
            {
                this.backgroundWorker11.RunWorkerAsync(); // 运行 backgroundWorker 组件

                JingDu form = new JingDu(this.backgroundWorker11, "刷新中");// 显示进度条窗体
                form.ShowDialog(this);
                form.Close();
                list2 = cal1.selectGongHuoFang();
                DataTable dt = new DataTable();
                dt.Columns.Add("Id22", typeof(int));
                dt.Columns.Add("PingMing11", typeof(String));
                dt.Columns.Add("HuoHao11", typeof(String));
                dt.Columns.Add("Guige11", typeof(String));
                dt.Columns.Add("SeHao12", typeof(String));
                dt.Columns.Add("Yanse11", typeof(String));
                dt.Columns.Add("DanJia11", typeof(String));
                dt.Columns.Add("GongHuoFangA11", typeof(String));
                dt.Columns.Add("GongHuoFangB11", typeof(String));
                dt.Columns.Add("BeiZhu11", typeof(String));
                foreach (clsBuiness.GongHuoFang s in list2)
                {
                    dt.Rows.Add(s.Id,s.PingMing, s.HuoHao, s.Guige, s.SeHao, s.Yanse, s.DanJia, s.GongHuoFangA , s.GongHuoFangB, s.BeiZhu);
                }
                dataGridView2.DataSource = dt;
                MessageBox.Show("刷新成功");


            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        #endregion

        #region 提交按钮

        private void toolStripLabel1_Click_1(object sender, EventArgs e)
        {
            try 
            {
                this.backgroundWorker11.RunWorkerAsync(); // 运行 backgroundWorker 组件

                JingDu form = new JingDu(this.backgroundWorker11, "提交中");// 显示进度条窗体
                form.ShowDialog(this);
                form.Close();
                DataTable dt = dataGridView2.DataSource as DataTable;
                cal1.insertGongHuoFang(dt);

                MessageBox.Show("提交成功！");
                toolStripLabel5_Click(sender, e);
            }
            catch (Exception ex) 
            {
                throw ex;
            }
            

        #endregion
        }

        private void toolStripLabel1_Click(object sender, EventArgs e)
        {
            DataTable dt = dataGridView2.DataSource as DataTable;
            cal1.insertGongHuoFang(dt);

            MessageBox.Show("提交成功！");
            toolStripLabel5_Click(sender, e);


        }


        private void toolStripLabel4_Click(object sender, EventArgs e)
        {
             {
            try 
            {
                this.backgroundWorker11.RunWorkerAsync(); // 运行 backgroundWorker 组件

                JingDu form = new JingDu(this.backgroundWorker11, "提交中");// 显示进度条窗体
                form.ShowDialog(this);
                form.Close();
                if (this.openFileDialog2.ShowDialog() == DialogResult.OK)
                {
                    string path = openFileDialog2.FileName;
                    if (!path.Equals(string.Empty))
                    {
                        list2 = cal1.readerGongHuoFangExcel(path);
                        DataTable dt = new DataTable();
                        dt.Columns.Add("Id22", typeof(int));
                        dt.Columns.Add("PingMing11", typeof(String));
                        dt.Columns.Add("HuoHao11", typeof(String));
                        dt.Columns.Add("Guige11", typeof(String));
                        dt.Columns.Add("SeHao12", typeof(String));
                        dt.Columns.Add("Yanse11", typeof(String));
                        dt.Columns.Add("DanJia11", typeof(String));
                        dt.Columns.Add("GongHuoFangA11", typeof(String));
                        dt.Columns.Add("GongHuoFangB11", typeof(String));
                        dt.Columns.Add("BeiZhu11", typeof(String));
                        foreach (clsBuiness.GongHuoFang s in list2)
                        {
                            dt.Rows.Add(s.Id, s.PingMing, s.HuoHao, s.Guige, s.SeHao, s.Yanse, s.DanJia, s.GongHuoFangA,s.GongHuoFangB,s.BeiZhu);
                        }
                        dataGridView2.DataSource = dt;
                    }
                    MessageBox.Show("读取成功！");

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            
             }}

   
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
        }

       
        }
    

