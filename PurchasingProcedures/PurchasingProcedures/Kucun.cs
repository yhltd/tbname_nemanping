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
    public partial class Kucun : Form
    {
        protected clsAllnewLogic cal;
        public Kucun()
        {
            InitializeComponent();
            cal = new clsAllnewLogic();
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
        }
        public void bindDatagridview() 
        {
            List<KuCun> list = cal.SelectKC();
            DataTable dt = new DataTable();
            dt.Columns.Add("Id", typeof(int));
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("Id"))
                {
                    dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                }
            }
            foreach (KuCun s in list)
            {
                dt.Rows.Add(s.Id, s.PingMing, s.HuoHao, s.SeHao, s.ShuLiang, s.GongHuoFang, s.CunFangDI);
            }
            dataGridView1.DataSource = dt;
        }
        private void toolStripLabel1_Click(object sender, EventArgs e)
        {
            this.backgroundWorker1.RunWorkerAsync();
            JingDu form = new JingDu(this.backgroundWorker1, "刷新中");// 显示进度条窗体
            form.ShowDialog(this);
            form.Close();
            bindDatagridview();
            MessageBox.Show("刷新成功！");
        }

        private void toolStripLabel2_Click(object sender, EventArgs e)
        {
            this.backgroundWorker1.RunWorkerAsync();
            JingDu form = new JingDu(this.backgroundWorker1, "提交中");// 显示进度条窗体
            form.ShowDialog(this);
            form.Close();
            DataTable dt = dataGridView1.DataSource as DataTable;
            if (dt == null)
            {
                dt = new DataTable();
                dt.Columns.Add("Id", typeof(int));
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("Id"))
                    {
                        dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                    }
                }
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1.Rows[i].Cells[1].Value != null)
                    {
                        dt.Rows.Add(dataGridView1.Rows[i].Cells[0].Value, dataGridView1.Rows[i].Cells[1].Value, dataGridView1.Rows[i].Cells[2].Value, dataGridView1.Rows[i].Cells[3].Value, dataGridView1.Rows[i].Cells[4].Value, dataGridView1.Rows[i].Cells[5].Value, dataGridView1.Rows[i].Cells[6].Value);
                    }
                }
            }
            cal.insertKucun(dt);
            MessageBox.Show("提交成功");
            bindDatagridview();
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
        private void toolStripLabel3_Click(object sender, EventArgs e)
        {
            this.backgroundWorker1.RunWorkerAsync();
            JingDu form = new JingDu(this.backgroundWorker1, "读取中");// 显示进度条窗体
            form.ShowDialog(this);
            form.Close();
             DialogResult queren = MessageBox.Show("读取的'EXCEL文件'后缀必须为.Xlsx，否则读取失败！","系统提示！",MessageBoxButtons.YesNo);
             if (queren == DialogResult.Yes)
             {
                 if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
                 {
                     string path = openFileDialog1.FileName;
                     if (!path.Equals(string.Empty))
                     {
                         if (path.Trim().Contains("xlsx"))
                         {
                             List<KuCun> list = cal.readerKucunExcel(path);
                             DataTable dt = new DataTable();
                             dt.Columns.Add("Id", typeof(int));
                             for (int i = 0; i < dataGridView1.Columns.Count; i++)
                             {
                                 if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("Id"))
                                 {
                                     dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                                 }
                             }
                             foreach (KuCun s in list)
                             {
                                 dt.Rows.Add(s.Id, s.PingMing, s.HuoHao, s.SeHao, s.ShuLiang, s.GongHuoFang, s.CunFangDI);
                             }
                             dataGridView1.DataSource = dt;
                         }
                         else 
                         {
                             MessageBox.Show("读取失败！原因:读取文件后缀非'xlsx'");
                         }
                     }
                 }
             }
             MessageBox.Show("读取完成！");
            
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
                    }
                    else
                    {
                        idtrr.Add(Convert.ToInt32(dataGridView1.SelectedRows[i - 1].Cells[0].Value));

                    }
                }
                cal.deleteKucun(idtrr);
                this.backgroundWorker1.RunWorkerAsync();
                JingDu form = new JingDu(this.backgroundWorker1, "删除中");// 显示进度条窗体
                form.ShowDialog(this);
                form.Close();
                MessageBox.Show("删除成功！");
                bindDatagridview();
                //comboBox1_SelectedIndexChanged(sender, e);

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void Kucun_Load(object sender, EventArgs e)
        {

        }
    }
}
