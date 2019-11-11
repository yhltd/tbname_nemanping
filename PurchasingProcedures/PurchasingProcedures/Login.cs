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
using System.Threading;
namespace PurchasingProcedures
{
    public partial class Login : Form
    {
        public clsAllnewLogic cals ;
        // 后台执行控件
        private BackgroundWorker bgWorker;
        // 后台操作是否正常完成
        private bool blnBackGroundWorkIsOK = false;
        public Login()
        {
            cals = new clsAllnewLogic();
            InitializeComponent();
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;

        }
       
        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void btn_Login_Click(object sender, EventArgs e)
        {
            try
            {
                this.backgroundWorker1.RunWorkerAsync(); // 运行 backgroundWorker 组件

                JingDu form = new JingDu(this.backgroundWorker1, "后台验证中");// 显示进度条窗体
                form.ShowDialog(this);
                form.Close();
                bool Loginpd = cals.Login(txt_User.Text.Trim(), txt_pwd.Text.Trim());

                if (Loginpd)
                {
                    MessageBox.Show("登陆成功！");

                    this.DialogResult = System.Windows.Forms.DialogResult.OK;
                    this.Close();
                }
                else
                {
                    MessageBox.Show("登录失败！原因：账号或密码错误");
                }
            }
            catch (Exception ex) 
            {
                MessageBox.Show("登录失败！原因：找不到账号或密码！");
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
    
    
        //你可以在这个方法内，实现你的调用，方法等。

        private void Login_Load(object sender, EventArgs e)
        {
           
        }

        private void backgroundWorker1_RunWorkerCompleted_1(object sender, RunWorkerCompletedEventArgs e)
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

        private void backgroundWorker1_DoWork_1(object sender, DoWorkEventArgs e)
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
    }
}
