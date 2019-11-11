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
    public partial class Factoryinput : Form
    {
        //protected DataTable dt;
        protected List<JiaGongChang> list1;
        protected Definefactoryinput cal1;
        public Factoryinput()
        {
            InitializeComponent();
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            cal1 = new Definefactoryinput();
            list1 = new List<JiaGongChang>();
        }

        #region 刷新按钮
        private void toolStripLabel5_Click(object sender, EventArgs e)
        {
            try
            {
                this.backgroundWorker1.RunWorkerAsync(); // 运行 backgroundWorker 组件

                JingDu form = new JingDu(this.backgroundWorker1, "刷新中");// 显示进度条窗体
                form.ShowDialog(this);
                form.Close();
                list1 = cal1.selectJiaGongChang();
                DataTable dt = new DataTable();
                dt.Columns.Add("id1", typeof(int));
                dt.Columns.Add("Name1", typeof(String));
                dt.Columns.Add("Address", typeof(String));
                dt.Columns.Add("Lianxiren", typeof(String));
                dt.Columns.Add("Phone", typeof(String));
                dt.Columns.Add("ZengZhiShui", typeof(String));
                dt.Columns.Add("Kaihuhang", typeof(String));
                dt.Columns.Add("Zhanghao", typeof(String));
                foreach (JiaGongChang s in list1)
                {
                    dt.Rows.Add(s.ID, s.Name, s.Address, s.Lianxiren, s.Phone, s.ZengZhiShui, s.Kaihuhang, s.Zhanghao);
                }
                dataGridView1.DataSource = dt;
                MessageBox.Show("刷新成功");


            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        #endregion


        #region 提交修改按钮
        private void toolStripLabel2_Click_1(object sender, EventArgs e)
        {
            try 
            {
                this.backgroundWorker1.RunWorkerAsync(); // 运行 backgroundWorker 组件

                JingDu form = new JingDu(this.backgroundWorker1, "提交中");// 显示进度条窗体
                form.ShowDialog(this);
                form.Close();
                DataTable dt = dataGridView1.DataSource as DataTable;
                cal1.insertJiaGongChang(dt);

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
            DataTable dt = dataGridView1.DataSource as DataTable;
            cal1.insertJiaGongChang(dt);

            MessageBox.Show("提交成功！");
            toolStripLabel5_Click(sender, e);


        }

        private void toolStripLabel4_Click(object sender, EventArgs e)
        {
            try 
            {
                this.backgroundWorker1.RunWorkerAsync(); // 运行 backgroundWorker 组件

                JingDu form = new JingDu(this.backgroundWorker1, "提交中");// 显示进度条窗体
                form.ShowDialog(this);
                form.Close();
                if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string path = openFileDialog1.FileName;
                    if (!path.Equals(string.Empty))
                    {
                        list1 = cal1.readerJiaGongChangExcel(path);
                        DataTable dt = new DataTable();
                        dt.Columns.Add("id1", typeof(int));
                        dt.Columns.Add("Name1", typeof(String));
                        dt.Columns.Add("Address", typeof(String));
                        dt.Columns.Add("Lianxiren", typeof(String));
                        dt.Columns.Add("Phone", typeof(String));
                        dt.Columns.Add("ZengZhiShui", typeof(String));
                        dt.Columns.Add("Kaihuhang", typeof(String));
                        dt.Columns.Add("Zhanghao", typeof(String));
                        foreach (JiaGongChang s in list1)
                        {
                            dt.Rows.Add(s.ID, s.Name, s.Address, s.Lianxiren, s.Phone, s.ZengZhiShui, s.Kaihuhang, s.Zhanghao);
                        }
                        dataGridView1.DataSource = dt;
                    }
                    MessageBox.Show("读取成功！");

                }
            }
            catch (Exception ex) 
            {
                throw ex;
            }
            
        }

        private void Factoryinput_Load(object sender, EventArgs e)
        {

        }

        private void toolStripLabel5_Click_1(object sender, EventArgs e)
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
    }
}