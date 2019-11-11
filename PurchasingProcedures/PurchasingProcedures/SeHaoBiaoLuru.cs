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
    public partial class SeHaoBiaoLuru : Form
    {
        //protected DataTable dt;
        protected List<Sehao> list;
        protected clsAllnewLogic cal ;
        public SeHaoBiaoLuru()
        {
            InitializeComponent();
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            cal = new clsAllnewLogic();
            list = new List<Sehao>();
        }



        #region 提交修改按钮
        private void toolStripLabel2_Click_1(object sender, EventArgs e)
        {
            this.backgroundWorker1.RunWorkerAsync(); // 运行 backgroundWorker 组件

                JingDu form = new JingDu(this.backgroundWorker1, "提交中");// 显示进度条窗体
                form.ShowDialog(this);
                form.Close();
                DataTable dt = dataGridView1.DataSource as DataTable;
                cal.insertSehao(dt);
            
            MessageBox.Show("提交成功！");
            toolStripLabel5_Click(sender, e);
        }
        #endregion
        #region 刷新按钮
        private void toolStripLabel5_Click(object sender, EventArgs e)
        {
            try
            {
                this.backgroundWorker1.RunWorkerAsync(); // 运行 backgroundWorker 组件

                JingDu form = new JingDu(this.backgroundWorker1, "刷新中");// 显示进度条窗体
                form.ShowDialog(this);
                form.Close();
                list = cal.selectSehao();
                DataTable dt = new DataTable();
                dt.Columns.Add("Id", typeof(int));
                dt.Columns.Add("Name", typeof(String));
                dt.Columns.Add("SeHao1", typeof(String));
                foreach (Sehao s in list) 
                {
                    dt.Rows.Add(s.Id, s.Name, s.SeHao1);
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

        private void toolStripLabel4_Click(object sender, EventArgs e)
        {
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
               string path = openFileDialog1.FileName;
               if (!path.Equals(string.Empty)) 
               {
                   this.backgroundWorker1.RunWorkerAsync(); // 运行 backgroundWorker 组件

                   JingDu form = new JingDu(this.backgroundWorker1, "读取中");// 显示进度条窗体
                   form.ShowDialog(this);
                   form.Close();
                   list = cal.readerSehaoExcel(path);
                   DataTable dt = new DataTable();
                   dt.Columns.Add("Id", typeof(int));
                   dt.Columns.Add("Name", typeof(String));
                   dt.Columns.Add("SeHao1", typeof(String));
                   foreach (Sehao s in list)
                   {
                       dt.Rows.Add(s.Id, s.Name, s.SeHao1);
                   }
                   dataGridView1.DataSource = dt;
               }
               MessageBox.Show("读取成功！");
               
            }
        }

        private void SeHaoBiaoLuru_Load(object sender, EventArgs e)
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
