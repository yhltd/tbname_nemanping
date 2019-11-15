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
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
        }
        public List<ChiMa_Dapeibiao> bindDataGirdview(string wheres) 
        {
            List<ChiMa_Dapeibiao> list = new List<ChiMa_Dapeibiao>();
            if (wheres.Equals(string.Empty))
            {
               list = cal.SelectChiMaDapei("");
            }
            else 
            {
                list = cal.SelectChiMaDapei(wheres);
            }
            return list;
            
        }
        private void toolStripLabel1_Click(object sender, EventArgs e)
        {
            
        }

        private void toolStripLabel2_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult dr = MessageBox.Show("确认要提交吗？","信息",MessageBoxButtons.YesNo);
                if (dr == DialogResult.Yes)
                {
                    this.backgroundWorker1.RunWorkerAsync(); // 运行 backgroundWorker 组件

                    JingDu form = new JingDu(this.backgroundWorker1, "提交中");// 显示进度条窗体
                    form.ShowDialog(this);
                    form.Close();
                    DataTable dt = dataGridView1.DataSource as DataTable;
                    if (dt == null)
                    {
                        dt = new DataTable();
                        dt.Columns.Add("id", typeof(int));
                        for (int i = 0; i < dataGridView1.Columns.Count; i++)
                        {
                            if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                            {
                                dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                            }
                        }
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            if (dataGridView1.Rows[i].Cells[6].Value != null)
                            {
                                dt.Rows.Add(dataGridView1.Rows[i].Cells[1].Value, dataGridView1.Rows[i].Cells[0].Value, dataGridView1.Rows[i].Cells[2].Value, dataGridView1.Rows[i].Cells[3].Value, dataGridView1.Rows[i].Cells[4].Value, dataGridView1.Rows[i].Cells[5].Value, dataGridView1.Rows[i].Cells[6].Value, dataGridView1.Rows[i].Cells[7].Value, dataGridView1.Rows[i].Cells[8].Value, dataGridView1.Rows[i].Cells[9].Value, dataGridView1.Rows[i].Cells[10].Value, dataGridView1.Rows[i].Cells[11].Value, dataGridView1.Rows[i].Cells[12].Value, dataGridView1.Rows[i].Cells[13].Value, dataGridView1.Rows[i].Cells[14].Value, dataGridView1.Rows[i].Cells[15].Value, dataGridView1.Rows[i].Cells[16].Value, dataGridView1.Rows[i].Cells[17].Value, dataGridView1.Rows[i].Cells[18].Value, dataGridView1.Rows[i].Cells[19].Value, dataGridView1.Rows[i].Cells[20].Value, dataGridView1.Rows[i].Cells[21].Value, dataGridView1.Rows[i].Cells[22].Value, dataGridView1.Rows[i].Cells[23].Value, dataGridView1.Rows[i].Cells[24].Value, dataGridView1.Rows[i].Cells[25].Value, dataGridView1.Rows[i].Cells[26].Value, dataGridView1.Rows[i].Cells[27].Value, dataGridView1.Rows[i].Cells[28].Value, dataGridView1.Rows[i].Cells[29].Value, dataGridView1.Rows[i].Cells[30].Value, dataGridView1.Rows[i].Cells[31].Value, dataGridView1.Rows[i].Cells[32].Value, dataGridView1.Rows[i].Cells[33].Value, dataGridView1.Rows[i].Cells[34].Value, dataGridView1.Rows[i].Cells[35].Value, dataGridView1.Rows[i].Cells[36].Value, dataGridView1.Rows[i].Cells[37].Value, dataGridView1.Rows[i].Cells[38].Value, dataGridView1.Rows[i].Cells[39].Value, dataGridView1.Rows[i].Cells[40].Value, dataGridView1.Rows[i].Cells[41].Value, dataGridView1.Rows[i].Cells[42].Value, dataGridView1.Rows[i].Cells[43].Value);
                            }
                        }
                    }
                    cal.InsertChima(dt, comboBox1.Text);

                    MessageBox.Show("提交成功！");
                    ChiMaDaPei_Load(sender, e);
                    bindDataGirdview(comboBox1.Text);
                }
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
                DialogResult queren = MessageBox.Show("读取的'EXCEL文件'后缀必须为.Xlsx，否则读取失败！", "系统提示！", MessageBoxButtons.YesNo);
                if (queren == DialogResult.Yes)
                {
                    List<ChiMa_Dapeibiao> list = new List<ChiMa_Dapeibiao>();
                    if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        string path = openFileDialog1.FileName;
                        if (!path.Equals(string.Empty))
                        {
                            if (path.Trim().Contains("xlsx"))
                            {
                                this.backgroundWorker1.RunWorkerAsync(); // 运行 backgroundWorker 组件

                                JingDu form = new JingDu(this.backgroundWorker1, "读取中");// 显示进度条窗体
                                form.ShowDialog(this);
                                form.Close();
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
                                MessageBox.Show("读取成功！");
                            }
                            else 
                            {
                                MessageBox.Show("读取失败！原因:读取文件后缀非'xlsx");
                            }
                        }
                    }
                   
                }
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
            try
            {
                DialogResult dr = MessageBox.Show("确定要删除选中信息？","信息",MessageBoxButtons.YesNo);
                if (dr == DialogResult.Yes)
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
                    cal.deleteChiMa(idtrr);
                    this.backgroundWorker1.RunWorkerAsync();
                    JingDu form = new JingDu(this.backgroundWorker1, "删除中");// 显示进度条窗体
                    form.ShowDialog(this);
                    form.Close();
                    MessageBox.Show("删除成功！");
                    bindDataGirdview(comboBox1.Text);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void ChiMaDaPei_Load(object sender, EventArgs e)
        {
            try
            {
                comboBox1.SelectedIndexChanged -= comboBox1_SelectedIndexChanged;
                List<ChiMa_Dapeibiao> list = bindDataGirdview("").GroupBy(p => p.BiaoName).Select(pc => pc.First()).ToList<ChiMa_Dapeibiao>();
                comboBox1.DisplayMember = "BiaoName";
                comboBox1.ValueMember = "id";
                comboBox1.DataSource = list;
                comboBox1.SelectedIndexChanged += comboBox1_SelectedIndexChanged;
            }
            catch (Exception ex) 
            {
                throw ex;
            }
            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {

                this.backgroundWorker1.RunWorkerAsync(); // 运行 backgroundWorker 组件

                JingDu form = new JingDu(this.backgroundWorker1, "刷新中");// 显示进度条窗体
                form.ShowDialog(this);
                form.Close();
                List<ChiMa_Dapeibiao> list = bindDataGirdview(comboBox1.Text);
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
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                if (comboBox1.Text.Equals(string.Empty))
                {
                    DataTable dt = new DataTable();
                    dt.Columns.Add("id", typeof(int));
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                        {
                            dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                        }
                    }
                    dataGridView1.DataSource = dt;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try 
            {
                DialogResult dr = MessageBox.Show("确定您要删除 名为'" + comboBox1.Text + "'的尺码搭配表吗？","信息",MessageBoxButtons.YesNo);
                if (dr == DialogResult.Yes)
                {
                    this.backgroundWorker1.RunWorkerAsync();
                    JingDu form = new JingDu(this.backgroundWorker1, "删除中");// 显示进度条窗体
                    form.ShowDialog(this);
                    form.Close();
                    cal.deleteChiMaBiao(comboBox1.Text);
                    MessageBox.Show("删除成功！");
                }
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }
    }
}
