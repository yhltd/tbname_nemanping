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
using System.Runtime.InteropServices;
using System.IO;
namespace PurchasingProcedures
{
    public partial class PeiSeBiaoLuru : Form
    {
        [DllImport("kernel32.dll")]
        public static extern IntPtr _lopen(string lpPathName, int iReadWrite);
        public const int OF_READWRITE = 2;
        public const int OF_SHARE_DENY_NONE = 0x40;
        public readonly IntPtr HFILE_ERROR = new IntPtr(-1);
        protected clsAllnewLogic cal = new clsAllnewLogic();
        [DllImport("kernel32.dll")]
        public static extern bool CloseHandle(IntPtr hObject);

        public PeiSeBiaoLuru()
        {
            InitializeComponent();
            //this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (cb_MianLiao.SelectedValue != "")
                {
                    List<PeiSe> list = cal.selectPeise(cb_MianLiao.Text);
                    DataTable dt = new DataTable();
                    dateTimePicker1.Value = Convert.ToDateTime(list[0].Date);
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                            dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                    }
                    foreach (PeiSe s in list)
                    {
                        dt.Rows.Add( s.PingMing, s.HuoHao, s.GuiGe, s.C61601C1, s.C61602C1, s.C61603C1, s.C61605C1, s.C61606C1, s.C61607C1, s.C61609C1, s.C61611C1, s.C61618C1, s.C61624C1, s.C61627C1, s.C61631C1, s.C61632C1, s.C61633C1, s.C61634C1, s.MianLiaoYanSe,s.Id,s.Fabrics,s.Date);
                    }
                    dataGridView1.DataSource = dt;
                }
            }
            catch (Exception ex) 
            {
                //throw ex;
                MessageBox.Show(ex.Message);

            }
            
        }

        private void PeiSeBiaoLuru_Load(object sender, EventArgs e)
        {
            try
            {
                List<PeiSe> list = cal.selectPeise("").GroupBy(p => new { p.Fabrics }).Select(pc =>pc.First()).ToList();
                cb_MianLiao.DisplayMember = "Fabrics";
                cb_MianLiao.ValueMember = "Id";
                cb_MianLiao.DataSource = list;
            }
            catch (Exception ex) 
            {
                //throw ex;
                MessageBox.Show(ex.Message);

            }
            
            
        }

        private void cb_MianLiao_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (cb_MianLiao.Text.Equals(string.Empty))
                {
                    DataTable dt = new DataTable();
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                    }
                    dataGridView1.DataSource = dt;
                }
            }
            catch (Exception ex) 
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void toolStripLabel2_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = dataGridView1.DataSource as DataTable;
                if (dt == null)
                {
                    dt = new DataTable();
                    //dt.Columns.Add("Id", typeof(int));
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        //if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("Id"))
                        //{
                            dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                        //}
                    }
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        if (dataGridView1.Rows[i].Cells[6].Value != null)
                        {
                            dt.Rows.Add(dataGridView1.Rows[i].Cells[0].Value, dataGridView1.Rows[i].Cells[1].Value, dataGridView1.Rows[i].Cells[2].Value, dataGridView1.Rows[i].Cells[3].Value, dataGridView1.Rows[i].Cells[4].Value, dataGridView1.Rows[i].Cells[5].Value, dataGridView1.Rows[i].Cells[6].Value, dataGridView1.Rows[i].Cells[7].Value, dataGridView1.Rows[i].Cells[8].Value, dataGridView1.Rows[i].Cells[9].Value, dataGridView1.Rows[i].Cells[10].Value, dataGridView1.Rows[i].Cells[11].Value, dataGridView1.Rows[i].Cells[12].Value, dataGridView1.Rows[i].Cells[13].Value, dataGridView1.Rows[i].Cells[14].Value, dataGridView1.Rows[i].Cells[15].Value, dataGridView1.Rows[i].Cells[16].Value, dataGridView1.Rows[i].Cells[17].Value, dataGridView1.Rows[i].Cells[18].Value, dataGridView1.Rows[i].Cells[19].Value, dataGridView1.Rows[i].Cells[20].Value, dataGridView1.Rows[i].Cells[21].Value);
                        }
                    }
                }
                this.backgroundWorker1.RunWorkerAsync(); // 运行 backgroundWorker 组件
                JingDu form = new JingDu(this.backgroundWorker1, "提交中");// 显示进度条窗体
                form.ShowDialog(this);
                form.Close();
                cal.insertPeise(dt, cb_MianLiao.Text, dateTimePicker1.Text);
                MessageBox.Show("提交成功！");
                string cbtext = cb_MianLiao.Text;
                PeiSeBiaoLuru_Load(sender,e);
                cb_MianLiao.Text = cbtext;
            }
            catch (Exception ex) 
            {
                //throw ex;
                MessageBox.Show(ex.Message);

            }
            
        }

        private void toolStripLabel3_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult queren = MessageBox.Show("读取的'EXCEL文件'后缀必须为.Xlsx，否则读取失败！","系统提示！",MessageBoxButtons.YesNo);
                if (queren == DialogResult.Yes)
                {
                    if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        string path = openFileDialog1.FileName;
                        if (!path.Equals(string.Empty))
                        {
                            if (!File.Exists(path))
                            {
                                MessageBox.Show("文件不存在！");
                                return;
                            }
                            IntPtr vHandle = _lopen(path, OF_READWRITE | OF_SHARE_DENY_NONE);
                            if (vHandle == HFILE_ERROR)
                            {
                                MessageBox.Show("文件被占用！");
                                return;
                            }
                            CloseHandle(vHandle);
                            if (path.Trim().Contains("xlsx"))
                            {
                                this.backgroundWorker1.RunWorkerAsync(); // 运行 backgroundWorker 组件
                                JingDu form = new JingDu(this.backgroundWorker1, "读取中");// 显示进度条窗体
                                form.ShowDialog(this);
                                form.Close();
                                List<PeiSe> list = cal.ReaderPeiSe(path);
                                DataTable dt = new DataTable();
                                if (list[0].Date != null)
                                {
                                    string strDate = DateTime.FromOADate(Convert.ToInt32(list[0].Date)).ToString("d");
                                    strDate = DateTime.Parse(strDate).ToString("dd-MMM-yyyy");
                                    dateTimePicker1.Value = Convert.ToDateTime(strDate);
                                }
                             
                                cb_MianLiao.DropDownStyle = ComboBoxStyle.DropDown;
                                cb_MianLiao.Text = list[0].Fabrics;
                                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                                {
                                    dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                                }
                                foreach (PeiSe s in list)
                                {
                                    if (s.PingMing != null)
                                    {
                                        dt.Rows.Add(s.PingMing, s.HuoHao, s.GuiGe, s.C61601C1, s.C61602C1, s.C61603C1, s.C61605C1, s.C61606C1, s.C61607C1, s.C61609C1, s.C61611C1, s.C61618C1, s.C61624C1, s.C61627C1, s.C61631C1, s.C61632C1, s.C61633C1, s.C61634C1, s.MianLiaoYanSe, s.Id, s.Fabrics, s.Date);
                                    }
                                }
                                dataGridView1.DataSource = dt;
                                MessageBox.Show("读取完成！");
                            }
                            else
                            {
                                MessageBox.Show("读取失败！原因:读取文件后缀非'xlsx'");
                            }
                            
                        }
                       
                        
                    }
                    
                }
            }
            catch (Exception ex) 
            {
                //throw ex;
                MessageBox.Show(ex.Message);
            }
            
        }

        private void 删除ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult dr = MessageBox.Show("确认要删除 该信息吗？", "系统提示", MessageBoxButtons.YesNo);
                if (dr == DialogResult.Yes)
                {
                    List<int> idtrr = new List<int>();
                    for (int i = this.dataGridView1.SelectedRows.Count; i > 0; i--)
                    {
                        if (dataGridView1.SelectedRows[i - 1].Cells[19].Value == null || dataGridView1.SelectedRows[i - 1].Cells[19].Value is DBNull)
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
                            idtrr.Add(Convert.ToInt32(dataGridView1.SelectedRows[i - 1].Cells[19].Value));

                        }
                    }

                    cal.deletePeiseSession(idtrr);
                    this.backgroundWorker1.RunWorkerAsync();
                    JingDu form = new JingDu(this.backgroundWorker1, "删除中");// 显示进度条窗体
                    form.ShowDialog(this);
                    form.Close();
                    MessageBox.Show("删除成功！");
                    comboBox1_SelectedIndexChanged(sender, e);
                }
            }
            catch (Exception ex)
            {
                //throw ex;
                MessageBox.Show(ex.Message);

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

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult dr = MessageBox.Show("确认要删除    面料号为：'" + cb_MianLiao.Text + "'的单耗表吗？", "提示", MessageBoxButtons.YesNo);
                if (dr == DialogResult.Yes)
                {
                    cal.deletps(cb_MianLiao.Text);
                    this.backgroundWorker1.RunWorkerAsync();
                    JingDu form = new JingDu(this.backgroundWorker1, "删除中");// 显示进度条窗体
                    form.ShowDialog(this);
                    form.Close();
                    MessageBox.Show("删除成功！");
                    PeiSeBiaoLuru_Load(sender, e);
                    cb_MianLiao.Text = cb_MianLiao.Text;
                    comboBox1_SelectedIndexChanged(sender, e);
                }
            }
            catch (Exception ex)
            {
                //throw ex;
                MessageBox.Show(ex.Message);

            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
