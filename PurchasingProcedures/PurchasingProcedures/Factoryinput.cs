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
    public partial class Factoryinput : Form
    {
        //protected DataTable dt;
        [DllImport("kernel32.dll")]
        public static extern IntPtr _lopen(string lpPathName, int iReadWrite);
        public const int OF_READWRITE = 2;
        public const int OF_SHARE_DENY_NONE = 0x40;
        [DllImport("kernel32.dll")]
        public static extern bool CloseHandle(IntPtr hObject);

        public readonly IntPtr HFILE_ERROR = new IntPtr(-1);

        protected List<JiaGongChang> list1;
        protected Definefactoryinput cal1;
        public Factoryinput()
        {
            InitializeComponent();
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            cal1 = new Definefactoryinput();
            list1 = new List<JiaGongChang>();
        }
        public void bindDataGirdview() 
        {
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
                dt.Rows.Add(s.id, s.Name, s.Address, s.Lianxiren, s.Phone, s.ZengZhiShui, s.Kaihuhang, s.Zhanghao);
            }
            dataGridView1.DataSource = dt;
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
                bindDataGirdview();
                MessageBox.Show("刷新成功");


            }
            catch (Exception ex)
            {
                //throw ex;
                MessageBox.Show(ex.Message);
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
                            dt.Rows.Add(dataGridView1.Rows[i].Cells[0].Value, dataGridView1.Rows[i].Cells[1].Value, dataGridView1.Rows[i].Cells[2].Value, dataGridView1.Rows[i].Cells[3].Value, dataGridView1.Rows[i].Cells[4].Value, dataGridView1.Rows[i].Cells[5].Value, dataGridView1.Rows[i].Cells[6].Value, dataGridView1.Rows[i].Cells[7].Value);
                        }
                    }
                }
                cal1.insertJiaGongChang(dt);

                MessageBox.Show("提交成功！");
                bindDataGirdview();
            }
            catch (Exception ex) 
            {
                //throw ex;
                MessageBox.Show(ex.Message);

            }
            

        #endregion
        }

        private void toolStripLabel1_Click(object sender, EventArgs e)
        {
            DataTable dt = dataGridView1.DataSource as DataTable;
            cal1.insertJiaGongChang(dt);

            MessageBox.Show("提交成功！");
            bindDataGirdview();

        }

        private void toolStripLabel4_Click(object sender, EventArgs e)
        {
            try 
            {
                DialogResult queren = MessageBox.Show("读取的'EXCEL文件'后缀必须为.Xlsx，否则读取失败！","系统提示！",MessageBoxButtons.YesNo);
                if (queren == DialogResult.Yes)
                {

                    this.backgroundWorker1.RunWorkerAsync(); // 运行 backgroundWorker 组件

                    JingDu form = new JingDu(this.backgroundWorker1, "读取中");// 显示进度条窗体
                    form.ShowDialog(this);
                    form.Close();
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
                                    dt.Rows.Add(s.id, s.Name, s.Address, s.Lianxiren, s.Phone, s.ZengZhiShui, s.Kaihuhang, s.Zhanghao);
                                }
                                dataGridView1.DataSource = dt;
                                MessageBox.Show("读取成功！");

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
                cal1.deleteJaGongChang(idtrr);
                this.backgroundWorker1.RunWorkerAsync();
                JingDu form = new JingDu(this.backgroundWorker1, "删除中");// 显示进度条窗体
                form.ShowDialog(this);
                form.Close();
                MessageBox.Show("删除成功！");
                bindDataGirdview();
                //comboBox1_SelectedIndexChanged(sender, e);

            }
            catch (Exception ex)
            {
                //throw ex;
                MessageBox.Show(ex.Message);

            }
        }
    }
}