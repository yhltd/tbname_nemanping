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
    public partial class GongHuoFang : Form
    {
        protected DataTable dt;
        protected List<clsBuiness.GongHuoFang> list2;
        protected Define1 cal1;
        [DllImport("kernel32.dll")]
        public static extern IntPtr _lopen(string lpPathName, int iReadWrite);
        public const int OF_READWRITE = 2;
        public const int OF_SHARE_DENY_NONE = 0x40;
        public readonly IntPtr HFILE_ERROR = new IntPtr(-1);
        [DllImport("kernel32.dll")]
        public static extern bool CloseHandle(IntPtr hObject);
        public GongHuoFang()
        {
            InitializeComponent();
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            //this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            cal1 = new Define1();
            list2 = new List<clsBuiness.GongHuoFang>();
        }
        public void bindGirdview() 
        {
            list2 = cal1.selectGongHuoFang();
            DataTable dt = new DataTable();
            dt.Columns.Add("Id22", typeof(int));
            dt.Columns.Add("PingMing11", typeof(String));
            dt.Columns.Add("HuoHao11", typeof(String));
            dt.Columns.Add("Guige11", typeof(String));
            dt.Columns.Add("SeHao12", typeof(String));
            dt.Columns.Add("Yanse11", typeof(String));
            dt.Columns.Add("DanJia11", typeof(String));
            dt.Columns.Add("GongHuoFangA11", typeof(String));
            dt.Columns.Add("GongHuoFangB11", typeof(String));
            dt.Columns.Add("BeiZhu11", typeof(String));
            foreach (clsBuiness.GongHuoFang s in list2)
            {
                dt.Rows.Add(s.Id, s.PingMing, s.HuoHao, s.Guige, s.SeHao, s.Yanse, s.DanJia, s.GongHuoFangA, s.GongHuoFangB, s.BeiZhu);
            }
            dataGridView2.DataSource = dt;
        }
        #region 刷新按钮
        private void toolStripLabel5_Click(object sender, EventArgs e)
        {
            try
            {
                this.backgroundWorker11.RunWorkerAsync(); // 运行 backgroundWorker 组件

                JingDu form = new JingDu(this.backgroundWorker11, "刷新中");// 显示进度条窗体
                form.ShowDialog(this);
                form.Close();
                bindGirdview();
                MessageBox.Show("刷新成功");


            }
            catch (Exception ex)
            {
                //throw ex;
                MessageBox.Show(ex.Message);
            }

        }
        #endregion

        #region 提交按钮

        private void toolStripLabel1_Click_1(object sender, EventArgs e)
        {
            try
            {
                this.backgroundWorker11.RunWorkerAsync(); // 运行 backgroundWorker 组件

                JingDu form = new JingDu(this.backgroundWorker11, "提交中");// 显示进度条窗体
                form.ShowDialog(this);
                form.Close();
                DataTable dt = dataGridView2.DataSource as DataTable;
                if (dt == null)
                {
                    dt = new DataTable();
                    dt.Columns.Add("Id", typeof(int));
                    for (int i = 0; i < dataGridView2.Columns.Count; i++)
                    {
                        if (!dataGridView2.Columns[i].HeaderCell.Value.ToString().Equals("Id"))
                        {
                            dt.Columns.Add(dataGridView2.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                        }
                    }
                    for (int i = 0; i < dataGridView2.Rows.Count; i++)
                    {
                        if (dataGridView2.Rows[i].Cells[1].Value != null)
                        {
                            dt.Rows.Add(dataGridView2.Rows[i].Cells[0].Value, dataGridView2.Rows[i].Cells[1].Value, dataGridView2.Rows[i].Cells[2].Value, dataGridView2.Rows[i].Cells[3].Value, dataGridView2.Rows[i].Cells[4].Value, dataGridView2.Rows[i].Cells[5].Value, dataGridView2.Rows[i].Cells[6].Value, dataGridView2.Rows[i].Cells[7].Value, dataGridView2.Rows[i].Cells[8].Value, dataGridView2.Rows[i].Cells[9].Value);
                        }
                    }
                }
                cal1.insertGongHuoFang(dt);

                MessageBox.Show("提交成功！");
                bindGirdview();
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

        }


        private void toolStripLabel4_Click(object sender, EventArgs e)
        {
            {
                try
                {
                    DialogResult queren = MessageBox.Show("读取的'EXCEL文件'后缀必须为.Xlsx，否则读取失败！", "系统提示！", MessageBoxButtons.YesNo);
                    if (queren == DialogResult.Yes)
                    {

                        this.backgroundWorker11.RunWorkerAsync(); // 运行 backgroundWorker 组件

                        JingDu form = new JingDu(this.backgroundWorker11, "提交中");// 显示进度条窗体
                        form.ShowDialog(this);
                        form.Close();
                        if (this.openFileDialog2.ShowDialog() == DialogResult.OK)
                        {
                            string path = openFileDialog2.FileName;
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

                                    list2 = cal1.readerGongHuoFangExcel(path);
                                    DataTable dt = new DataTable();
                                    dt.Columns.Add("Id22", typeof(int));
                                    dt.Columns.Add("PingMing11", typeof(String));
                                    dt.Columns.Add("HuoHao11", typeof(String));
                                    dt.Columns.Add("Guige11", typeof(String));
                                    dt.Columns.Add("SeHao12", typeof(String));
                                    dt.Columns.Add("Yanse11", typeof(String));
                                    dt.Columns.Add("DanJia11", typeof(String));
                                    dt.Columns.Add("GongHuoFangA11", typeof(String));
                                    dt.Columns.Add("GongHuoFangB11", typeof(String));
                                    dt.Columns.Add("BeiZhu11", typeof(String));
                                    foreach (clsBuiness.GongHuoFang s in list2)
                                    {
                                        dt.Rows.Add(s.Id, s.PingMing, s.HuoHao, s.Guige, s.SeHao, s.Yanse, s.DanJia, s.GongHuoFangA, s.GongHuoFangB, s.BeiZhu);
                                    }
                                    dataGridView2.DataSource = dt;
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

        private void GongHuoFang_Load(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            try
            {
                List<int> idtrr = new List<int>();
                for (int i = this.dataGridView2.SelectedRows.Count; i > 0; i--)
                {
                    if (dataGridView2.SelectedRows[i - 1].Cells[0].Value == null || dataGridView2.SelectedRows[i - 1].Cells[0].Value is DBNull)
                    {
                        DataRowView drv = dataGridView2.SelectedRows[i - 1].DataBoundItem as DataRowView;
                        if (drv != null)
                        {
                            drv.Delete();
                            i = i - 1;
                        }
                        i = i - 1;
                    }
                    else
                    {
                        idtrr.Add(Convert.ToInt32(dataGridView2.SelectedRows[i - 1].Cells[0].Value));

                    }
                }
                cal1.deleteGongHuoFang(idtrr);
                this.backgroundWorker11.RunWorkerAsync();
                JingDu form = new JingDu(this.backgroundWorker11, "删除中");// 显示进度条窗体
                form.ShowDialog(this);
                form.Close();
                MessageBox.Show("删除成功！");
                bindGirdview();
                //bindDataGirdview();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }
    }


}


