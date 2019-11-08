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
    public partial class KuanShibiao : Form
    {
        protected clsAllnewLogic cal = new clsAllnewLogic();
        public KuanShibiao()
        {
            InitializeComponent();
        }

        private void 刷新ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void KuanShibiao_Load(object sender, EventArgs e)
        {

        }

        private void toolStripLabel1_Click(object sender, EventArgs e)
        {
            try
            {
                this.backgroundWorker1.RunWorkerAsync(); // 运行 backgroundWorker 组件

                JingDu form = new JingDu(this.backgroundWorker1, "刷新中");// 显示进度条窗体
                form.ShowDialog(this);
                form.Close();
                List<KuanShiBiao> list = cal.SelectKuanshi();
                DataTable dt = new DataTable();
                dt.Columns.Add("Id", typeof(int));
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("Id"))
                    {
                        dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                    }
                }
                foreach (KuanShiBiao s in list)
                {
                    dt.Rows.Add(s.Id, s.STYLE, s.DESC, s.FABRIC, s.JACKET, s.PANT, s.mark1, s.ShuoMing, s.mark2);
                }
                dataGridView1.DataSource = dt;
                MessageBox.Show("刷新成功！");

            }
            catch (Exception ex) 
            {
                throw ex;
            }
        }

        private void toolStripLabel2_Click(object sender, EventArgs e)
        {
            try
            {
                this.backgroundWorker1.RunWorkerAsync(); // 运行 backgroundWorker 组件

                JingDu form = new JingDu(this.backgroundWorker1, "提交中");// 显示进度条窗体
                form.ShowDialog(this);
                form.Close();
                DataTable dt = dataGridView1.DataSource as DataTable;
                cal.insertKuanShi(dt);
                MessageBox.Show("提交成功！");
            }
            catch (Exception ex)
            {
                throw ex;
            }
            
        }

        private void toolStripLabel3_Click(object sender, EventArgs e)
        {
            try 
            {
                List<KuanShiBiao> list = new List<KuanShiBiao>();
                if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string path = openFileDialog1.FileName;
                    if (!path.Equals(string.Empty))
                    {
                        this.backgroundWorker1.RunWorkerAsync(); // 运行 backgroundWorker 组件

                        JingDu form = new JingDu(this.backgroundWorker1, "读取中");// 显示进度条窗体
                        form.ShowDialog(this);
                        form.Close();
                        list = cal.readerKuanshi(path);
                        DataTable dt = new DataTable();
                        dt.Columns.Add("Id", typeof(int));
                        for (int i = 0; i < dataGridView1.Columns.Count; i++)
                        {
                            if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("Id"))
                            {
                                dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                            }
                        }
                        foreach (KuanShiBiao s in list)
                        {
                            dt.Rows.Add(s.Id, s.STYLE, s.DESC, s.FABRIC, s.JACKET, s.PANT, s.mark1, s.ShuoMing, s.mark2);
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
