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

namespace PurchasingProcedures
{
    public partial class JingDu : Form
    {
        private BackgroundWorker backgroundWorker1; //ProcessForm 窗体事件(进度条窗体)
        public JingDu(BackgroundWorker backgroundWorker1,string str)
        {
            InitializeComponent();
            label1.Text = str;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;

            this.backgroundWorker1 = backgroundWorker1;
            this.backgroundWorker1.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
            this.backgroundWorker1.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker1_RunWorkerCompleted);
        }
        void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.Close();//执行完之后，直接关闭页面
        }

        void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.progressBar1.Value = e.ProgressPercentage;
        }

        private void JingDu_Load(object sender, EventArgs e)
        {

        }

       
    }
}
