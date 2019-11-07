﻿using System;
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
            bool Loginpd = cals.Login(txt_User.Text.Trim(),txt_pwd.Text.Trim());
            if (blnBackGroundWorkIsOK) 
            {
                MessageBox.Show("登陆成功！");

                this.DialogResult = System.Windows.Forms.DialogResult.OK;
                this.Close();
            }
            else
            {
                MessageBox.Show("登录失败！");
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
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            for (int i = 0; i < 100; i++)
            {
                Thread.Sleep(100);
                worker.ReportProgress(i);
                if (worker.CancellationPending)  // 如果用户取消则跳出处理数据代码 
                {
                    e.Cancel = true;
                    break;
                }
            }
        }
        private void Login_Load(object sender, EventArgs e)
        {
           ;
        }
    }
}
