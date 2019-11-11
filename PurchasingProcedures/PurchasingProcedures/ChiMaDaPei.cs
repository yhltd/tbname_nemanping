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
    public partial class ChiMaDaPei : Form
    {
        public clsAllnewLogic cal = new clsAllnewLogic();
        public ChiMaDaPei()
        {
            InitializeComponent();
        }
        
        private void toolStripLabel1_Click(object sender, EventArgs e)
        {
            try
            {
                this.backgroundWorker1.RunWorkerAsync(); // 运行 backgroundWorker 组件

                JingDu form = new JingDu(this.backgroundWorker1, "刷新中");// 显示进度条窗体
                form.ShowDialog(this);
                form.Close();
                List<ChiMa_Dapeibiao> list =cal.SelectChiMaDapei();
                DataTable dt = new DataTable();
                dt.Columns.Add("id", typeof(int));
                for(int i =0 ;i < dataGridView1.Columns.Count;i++)
                {
                    if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                    {
                        dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                    }
                }
               foreach (ChiMa_Dapeibiao s in list) 
                {
                    dt.Rows.Add(s.id,s.LOT__面料,s.STYLE_款式,s.ART_货号,s.COLOR_颜色,s.COLOR__颜色编号,s.JACKET_上衣_PANT_裤子,s.C34R_28,s.C36R_30,s.C38R_32,s.C40R___34,s.C42R_36,s.C44R_38,s.C46R_40,s.C48R_42,s.C50R_44,s.C52R_46,s.C54R_48,s.C56R_50,s.C58R_52,s.C60R_54,s.C62R_56,s.C36L_30,s.C38L_32,s.C40L_34,s.C42L_36,s.C44L_38,s.C46L_40,s.C48L_42,s.C50L_44,s.C52L_46,s.C54L_48,s.C56L_50,s.C58L_52,s.C60L_54,s.C62L_56,s.C34S_28,s.C36S_30,s.C38S_32,s.C40S_34,s.C42S_36,s.C44S_38,s.C46S_40,s.DingdanHeji);
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
                cal.InsertChima(dt);

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
                this.backgroundWorker1.RunWorkerAsync(); // 运行 backgroundWorker 组件

                JingDu form = new JingDu(this.backgroundWorker1, "读取中");// 显示进度条窗体
                form.ShowDialog(this);
                form.Close();
                List<ChiMa_Dapeibiao> list = new List<ChiMa_Dapeibiao>();
                if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string path = openFileDialog1.FileName;
                    if (!path.Equals(string.Empty))
                    {
                        list = cal.ReaderChiMaDapei(path);
                        DataTable dt = new DataTable();
                        dt.Columns.Add("id", typeof(int));
                        for (int i = 0; i < dataGridView1.Columns.Count; i++)
                        {
                            if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                            {
                                dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                            }
                        }
                        foreach (ChiMa_Dapeibiao s in list)
                        {
                            dt.Rows.Add(s.id, s.LOT__面料, s.STYLE_款式, s.ART_货号, s.COLOR_颜色, s.COLOR__颜色编号, s.JACKET_上衣_PANT_裤子, s.C34R_28, s.C36R_30, s.C38R_32, s.C40R___34, s.C42R_36, s.C44R_38, s.C46R_40, s.C48R_42, s.C50R_44, s.C52R_46, s.C54R_48, s.C56R_50, s.C58R_52, s.C60R_54, s.C62R_56, s.C36L_30, s.C38L_32, s.C40L_34, s.C42L_36, s.C44L_38, s.C46L_40, s.C48L_42, s.C50L_44, s.C52L_46, s.C54L_48, s.C56L_50, s.C58L_52, s.C60L_54, s.C62L_56, s.C34S_28, s.C36S_30, s.C38S_32, s.C40S_34, s.C42S_36, s.C44S_38, s.C46S_40, s.DingdanHeji);
                        }
                        dataGridView1.DataSource = dt;
                    }
                }
                MessageBox.Show("读取成功！");
                
            }
            catch (Exception ex) 
            {
                throw ex;
            }
        }

        private void groupBox1_Enter(object sender, EventArgs e)
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

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void 删除ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}
