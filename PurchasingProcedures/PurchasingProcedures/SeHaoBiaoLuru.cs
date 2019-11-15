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
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            cal = new clsAllnewLogic();
            list = new List<Sehao>();
        }



        #region 提交修改按钮
        private void toolStripLabel2_Click_1(object sender, EventArgs e)
        {
                this.backgroundWorker1.RunWorkerAsync(); // 运行 backgroundWorker 组件
                DataTable dt =dataGridView1 .DataSource as DataTable;
                

                if (dt == null)
                {
                    dt = new DataTable();
                    dt.Columns.Add("Id", typeof(int));
                    dt.Columns.Add("Name", typeof(String));
                    dt.Columns.Add("SeHao1", typeof(String));
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        if (dataGridView1.Rows[i].Cells[0].Value != null)
                        {
                            if (!dataGridView1.Rows[i].Cells[0].Value.Equals(string.Empty))
                            {
                                dt.Rows.Add(dataGridView1.Rows[i].Cells[2].Value, dataGridView1.Rows[i].Cells[1].Value, dataGridView1.Rows[i].Cells[0].Value);
                            }
                        }
                    }
                }
                cal.insertSehao(dt);
                JingDu form = new JingDu(this.backgroundWorker1, "提交中");// 显示进度条窗体
                form.ShowDialog(this);
                form.Close();
                
            
            MessageBox.Show("提交成功！");
            bindDatagridView();
        }
        #endregion
        #region 绑定datagridview
        public void bindDatagridView() 
        {
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
        }
        #endregion
        #region 刷新按钮
        private void toolStripLabel5_Click(object sender, EventArgs e)
        {
            try
            {
                bindDatagridView();
                this.backgroundWorker1.RunWorkerAsync(); // 运行 backgroundWorker 组件
                JingDu form = new JingDu(this.backgroundWorker1, "刷新中");// 显示进度条窗体
                form.ShowDialog(this);
                form.Close();
                
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
                   list = cal.readerSehaoExcel(path);
                   this.backgroundWorker1.RunWorkerAsync(); // 运行 backgroundWorker 组件
                   JingDu form = new JingDu(this.backgroundWorker1, "读取中");// 显示进度条窗体
                   form.ShowDialog(this);
                   form.Close();
                   
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

        private void 删除ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try 
            {
                List<int> idtrr = new List<int>();
                for (int i = this.dataGridView1.SelectedRows.Count; i > 0; i--)
                {
                    if (dataGridView1.SelectedRows[i - 1].Cells[0].Value == null || dataGridView1.SelectedRows[i - 1].Cells[0].Value is DBNull)
                    {
                        DataRowView drv = dataGridView1.SelectedRows[i - 1].DataBoundItem as DataRowView;
                        if (drv != null)
                        {
                            drv.Delete();
                            i = i - 1;
                        }
                        i = i - 1;
                    }
                    else 
                    {
                        idtrr.Add(Convert.ToInt32(dataGridView1.SelectedRows[i - 1].Cells[0].Value));

                    }
                }
                cal.deleteSehao(idtrr);
                this.backgroundWorker1.RunWorkerAsync();
                JingDu form = new JingDu(this.backgroundWorker1, "删除中");// 显示进度条窗体
                form.ShowDialog(this);
                form.Close();
                MessageBox.Show("删除成功！");
                bindDatagridView();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void toolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }


        

   
    }
}
