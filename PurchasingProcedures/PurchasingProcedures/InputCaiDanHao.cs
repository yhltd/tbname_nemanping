using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using clsBuiness;
using logic;
namespace PurchasingProcedures
{
    public partial class InputCaiDanHao : Form
    {
        public clsAllnewLogic cal = new clsAllnewLogic();
        public string insertType;
        public string frmLabel;
        public InputCaiDanHao(string type,string label)
        {
            InitializeComponent();
            insertType = type;
            frmLabel = label;
             this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.backgroundWorker1.RunWorkerAsync();
            JingDu frm = new JingDu(this.backgroundWorker1, "生成裁单表中....");
            frm.ShowDialog();
            frm.Close();
            if(insertType.Equals("裁单"))
            {
                CaiDan C = new CaiDan(textBox1.Text,comboBox1.Text);
                if (!C.IsDisposed)
                {
                    C.ShowDialog();
                }
                this.Close();
            }
        }

        private void InputCaiDanHao_Load(object sender, EventArgs e)
        {
            label1.Text = frmLabel;
            List<ChiMa_Dapeibiao> list = cal.SelectChiMaDapei("").GroupBy(p => p.BiaoName).Select(pc => pc.First()).ToList<ChiMa_Dapeibiao>();
            comboBox1.DisplayMember = "BiaoName";
            comboBox1.ValueMember = "id";
            comboBox1.DataSource = list;
            List<clsBuiness.KuanShiBiao> cdlist = cal.SelectKuanshi().GroupBy(g => g.STYLE).Select(pc=>pc.First()).ToList<clsBuiness.KuanShiBiao>();
            textBox1.DataSource = cdlist;
            textBox1.DisplayMember = "STYLE";
            textBox1.ValueMember = "Id";

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
