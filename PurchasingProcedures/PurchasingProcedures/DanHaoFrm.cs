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
using System.Runtime.InteropServices;
using System.IO;
namespace PurchasingProcedures
{
    public partial class DanHaoFrm : Form
    {
        protected clsAllnewLogic cal = new clsAllnewLogic();
        [DllImport("kernel32.dll")]
        public static extern IntPtr _lopen(string lpPathName, int iReadWrite);
        public const int OF_READWRITE = 2;
        public const int OF_SHARE_DENY_NONE = 0x40;
        public readonly IntPtr HFILE_ERROR = new IntPtr(-1);
        [DllImport("kernel32.dll")]
        public static extern bool CloseHandle(IntPtr hObject);
        public DanHaoFrm()
        {
            InitializeComponent();
            //this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
        }

        private void toolStripLabel1_Click(object sender, EventArgs e)
        {
          
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (comboBox1.SelectedValue != "")
                {
                    List<DanHao> list = cal.SelectDanHao(comboBox1.Text.ToString());
                    DataTable dt1 = new DataTable();
                    dt1.Columns.Add("Id", typeof(int));
                    DataTable dt2 = new DataTable();
                    dt2.Columns.Add("Id", typeof(int));
                    txt_STYLE.Text = list[0].Style;
                    dateTimePicker1.Text = list[0].DATE;
                    txt_mfcf.Text = list[0].FABRIC_CONTENT;
                    txt_JGC.Text = list[0].JiaGongChang;
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("Id"))
                        {
                            dt1.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                            dt2.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                        }
                    }
                    foreach (DanHao s in list)
                    {
                        if (s.Type != null)
                        {
                            if (s.Type.Equals("面料"))
                            {
                                dt1.Rows.Add(s.Id, s.CaiDanNo, s.Style, s.FABRIC_CONTENT, s.DATE, s.JiaGongChang, s.Name, s.HuoHao, s.GuiGe, s.Yanse, s.Danjia, s.DanHao1, s.Jine, s.BeiZhu, s.ChangShang, s.Type);
                            }
                            else
                            {
                                dt2.Rows.Add(s.Id, s.CaiDanNo, s.Style, s.FABRIC_CONTENT, s.DATE, s.JiaGongChang, s.Name, s.HuoHao, s.GuiGe, s.Yanse, s.Danjia, s.DanHao1, s.Jine, s.BeiZhu, s.ChangShang, s.Type);
                            }
                        }
                    }
                    dataGridView1.DataSource = dt1;
                    dataGridView2.DataSource = dt2;
                    //MessageBox.Show("刷新成功！");
                }

            }
            catch (Exception ex)
            {
                //throw ex;
                MessageBox.Show(ex.Message);
            }
        }

        private void DanHaoFrm_Load(object sender, EventArgs e)
        {
            try
            {
                comboBox1.SelectedIndexChanged -= comboBox1_SelectedIndexChanged;
                //List<DanHao> list = cal.SelectDanHao("").GroupBy(d =>d.CaiDanNo).Select(p => p.First).ToList<DanHao>;
                List<DanHao> list = cal.SelectDanHao("").GroupBy(d => new { d.CaiDanNo }).Select(s => s.First()).ToList<DanHao>();
                comboBox1.DisplayMember = "CaiDanNo";
                comboBox1.ValueMember = "Id";
                comboBox1.DataSource = list;
                if (list != null && list.Count > 0 )
                {
                    comboBox1.SelectedIndex = 0;//设置默认值
                }
                comboBox1.SelectedIndexChanged += comboBox1_SelectedIndexChanged;//注册事件
            }catch(Exception ex)
            {
                //throw ex;
                MessageBox.Show(ex.Message);
            }
        }

        private void toolStripLabel2_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult queren = MessageBox.Show("确认提交？", "系统提示", MessageBoxButtons.YesNo);
                if (queren == DialogResult.Yes)
                {

                        DataTable newDataTable = new DataTable();
                        newDataTable.Columns.Add("Id", typeof(int));
                        for (int i = 0; i < dataGridView1.Columns.Count; i++)
                        {
                            if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("Id"))
                            {
                                newDataTable.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                            }
                        }
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            if (dataGridView1.Rows[i].Cells[6].Value!=null )
                            {
                                newDataTable.Rows.Add(dataGridView1.Rows[i].Cells[0].Value, dataGridView1.Rows[i].Cells[1].Value, dataGridView1.Rows[i].Cells[2].Value, dataGridView1.Rows[i].Cells[3].Value, dataGridView1.Rows[i].Cells[4].Value, dataGridView1.Rows[i].Cells[5].Value, dataGridView1.Rows[i].Cells[6].Value, dataGridView1.Rows[i].Cells[7].Value, dataGridView1.Rows[i].Cells[8].Value, dataGridView1.Rows[i].Cells[9].Value, dataGridView1.Rows[i].Cells[10].Value, dataGridView1.Rows[i].Cells[11].Value, dataGridView1.Rows[i].Cells[12].Value, dataGridView1.Rows[i].Cells[13].Value, dataGridView1.Rows[i].Cells[14].Value, dataGridView1.Rows[i].Cells[15].Value, "面料");
                            }
                        }

                        for (int i = 0; i < dataGridView2.Rows.Count; i++)
                        {
                            if (dataGridView2.Rows[i].Cells[6].Value!=null)
                            {
                                newDataTable.Rows.Add(dataGridView2.Rows[i].Cells[0].Value, dataGridView2.Rows[i].Cells[1].Value, dataGridView2.Rows[i].Cells[2].Value, dataGridView2.Rows[i].Cells[3].Value, dataGridView2.Rows[i].Cells[4].Value, dataGridView2.Rows[i].Cells[5].Value, dataGridView2.Rows[i].Cells[6].Value, dataGridView2.Rows[i].Cells[7].Value, dataGridView2.Rows[i].Cells[8].Value, dataGridView2.Rows[i].Cells[9].Value, dataGridView2.Rows[i].Cells[10].Value, dataGridView2.Rows[i].Cells[11].Value, dataGridView2.Rows[i].Cells[12].Value, dataGridView2.Rows[i].Cells[13].Value, dataGridView2.Rows[i].Cells[14].Value, dataGridView2.Rows[i].Cells[15].Value, "辅料");
                            }
                       }
                        foreach (DataRow dr in newDataTable.Rows)
                        {
                            //if (dr[1] is DBNull || dr[1].Equals(string.Empty))
                            //{
                                dr[1] = comboBox1.Text;
                            //}
                            //if (dr[2] is DBNull || dr[2].Equals(string.Empty))
                            //{
                                dr[2] = txt_STYLE.Text;
                            //}
                            //if (dr[3] is DBNull || dr[3].Equals(string.Empty))
                            //{
                                dr[3] = txt_mfcf.Text;
                            //}
                            //if (dr[4] is DBNull || dr[4].Equals(string.Empty))
                            //{
                                dr[4] = dateTimePicker1.Text;
                            //}
                            //if (dr[5] is DBNull || dr[5].Equals(string.Empty))
                            //{
                                dr[5] = txt_JGC.Text;
                            //}
                        }
                        this.backgroundWorker1.RunWorkerAsync(); // 运行 backgroundWorker 组件
                        JingDu form = new JingDu(this.backgroundWorker1, "提交中");// 显示进度条窗体
                        form.ShowDialog(this);
                        form.Close();
                        cal.insertDanhao(newDataTable);
                        MessageBox.Show("提交成功！");
                        string comboboxtext = comboBox1.Text;
                        DanHaoFrm_Load(sender, e);
                        comboBox1.Text = comboboxtext;
                }
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
                                List<DanHao> list = cal.Readerdh(path);
                                this.backgroundWorker1.RunWorkerAsync(); // 运行 backgroundWorker 组件
                                JingDu form = new JingDu(this.backgroundWorker1, "读取中");// 显示进度条窗体
                                form.ShowDialog(this);
                                form.Close();
                                DataTable dt1 = new DataTable();
                                dt1.Columns.Add("Id", typeof(int));
                                DataTable dt2 = new DataTable();
                                dt2.Columns.Add("Id", typeof(int));
                                txt_STYLE.Text = list[0].Style;
                                dateTimePicker1.Text = list[0].DATE;
                                txt_mfcf.Text = list[0].FABRIC_CONTENT;
                                txt_JGC.Text = list[0].JiaGongChang;
                                this.comboBox1.DropDownStyle = ComboBoxStyle.DropDown;
                                this.comboBox1.Text = list[0].CaiDanNo;
                                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                                {
                                    if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("Id"))
                                    {
                                        dt1.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                                        dt2.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                                    }
                                }
                                foreach (DanHao s in list)
                                {
                                    if (s.Type != null)
                                    {
                                        if (!s.Name.Contains("金额") && !s.Name.Equals("面料") && !s.Name.Contains("辅料") && !s.Name.Equals(string.Empty)) 
                                        {
                                            if (s.Type.Equals("面料"))
                                            {
                                                dt1.Rows.Add(s.Id, s.CaiDanNo, s.Style, s.FABRIC_CONTENT, s.DATE, s.JiaGongChang, s.Name, s.HuoHao, s.GuiGe, s.Yanse, s.Danjia, s.DanHao1, s.Jine, s.BeiZhu, s.ChangShang, s.Type);
                                            }
                                            else
                                            {
                                                dt2.Rows.Add(s.Id, s.CaiDanNo, s.Style, s.FABRIC_CONTENT, s.DATE, s.JiaGongChang, s.Name, s.HuoHao, s.GuiGe, s.Yanse, s.Danjia, s.DanHao1, s.Jine, s.BeiZhu, s.ChangShang, s.Type);
                                            }

                                        }
                                    }
                                }
                                dataGridView1.DataSource = dt1;
                                dataGridView2.DataSource = dt2;
                                Jisuan();
                            }
                            else
                            {
                                MessageBox.Show("读取失败！原因:读取文件后缀非'xlsx'");
                            }
                        }
                    }
                }
            }
            catch(Exception ex)
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

        private void 删除ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult dr = MessageBox.Show("确认要删除选中信息吗？", "提示", MessageBoxButtons.YesNo);
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
                        }
                        else
                        {
                            idtrr.Add(Convert.ToInt32(dataGridView1.SelectedRows[i - 1].Cells[0].Value));

                        }
                    }
                    for (int i = this.dataGridView2.SelectedRows.Count; i > 0; i--)
                    {
                        if (dataGridView2.SelectedRows[i - 1].Cells[0].Value == null || dataGridView2.SelectedRows[i - 1].Cells[0].Value is DBNull)
                        {
                            DataRowView drv = dataGridView2.SelectedRows[i - 1].DataBoundItem as DataRowView;
                            drv.Delete();
                            i = i - 1;
                        }
                        else
                        {
                            idtrr.Add(Convert.ToInt32(dataGridView2.SelectedRows[i - 1].Cells[0].Value));

                        }
                    }
                    cal.deleteDanHao(idtrr);
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

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult dr = MessageBox.Show("确认要删除    裁单号为：'" + comboBox1.Text + "'的单耗表吗？", "提示", MessageBoxButtons.YesNo);
                if (dr == DialogResult.Yes)
                {
                    cal.deletDh(comboBox1.Text);
                    this.backgroundWorker1.RunWorkerAsync();
                    JingDu form = new JingDu(this.backgroundWorker1, "删除中");// 显示进度条窗体
                    form.ShowDialog(this);
                    form.Close();
                    MessageBox.Show("删除成功！");
                    DanHaoFrm_Load(sender, e);
                    comboBox1.Text = comboBox1.Text;
                    comboBox1_SelectedIndexChanged(sender, e);
                }
            }
            catch (Exception ex) 
            {
                //throw ex;
                MessageBox.Show(ex.Message);
            }
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            Jisuan();
        }

        private void Jisuan()
        {
            double flSum = 0;
            double mlpj = 0;
            //for (int i = 0; i < dataGridView1.Rows.Count; i++)
            //{
            //    if (dataGridView1.Rows[i].Cells[6].Value != null)
            //    {
            //    }
            //}

            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {

                if (dataGridView2.Rows[i].Cells[6].Value != null && dataGridView2.Rows[i].Cells[12].Value!=null &&  cal.IsNumber(dataGridView2.Rows[i].Cells[12].Value.ToString()))
                {
                    flSum = flSum + Convert.ToDouble(dataGridView2.Rows[i].Cells[12].Value);
                }
            }
            FL_Sum.Text ="辅料总和："+ flSum.ToString();
        }
    }
}
