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
    public partial class inputMianLiaoDingGou : Form
    {
        protected clsAllnewLogic cal;
        protected Definefactoryinput dfi;
        public inputMianLiaoDingGou()
        {
            
            InitializeComponent();
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            cal = new clsAllnewLogic();
            dfi = new Definefactoryinput();
            
        }

        private void inputMianLiaoDingGou_Load(object sender, EventArgs e)
        {
            try
            {

                List<JiaGongChang> jgc = dfi.selectJiaGongChang().GroupBy(j => j.Name).Select(s => s.First()).ToList<JiaGongChang>();
                cb_jgc.DataSource = jgc;
                cb_jgc.DisplayMember = "Name";
                cb_jgc.ValueMember = "id";

            }
            catch (Exception ex) { throw ex; }
            
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
            if (!txt_ml.Text.Equals(string.Empty) && !txt_ks.Text.Equals(string.Empty) && !cb_jgc.Text.Equals(string.Empty))
            {
                this.backgroundWorker1.RunWorkerAsync();
                JingDu form = new JingDu(this.backgroundWorker1, "生成核算表中");// 显示进度条窗体
                form.ShowDialog(this);
                form.Close();
                MianFuLiaoDingGou mfdg = new MianFuLiaoDingGou(txt_ml.Text,txt_ks.Text,cb_jgc.Text);
                mfdg.ShowDialog();
                this.Close();
            }
            else 
            {
                MessageBox.Show("生成失败！原因:数据不能为空");
            }
        }
    }
}
