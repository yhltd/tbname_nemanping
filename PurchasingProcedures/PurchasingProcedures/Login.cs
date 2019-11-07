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
namespace PurchasingProcedures
{
    public partial class Login : Form
    {
        public clsAllnewLogic cals ;
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
            if (Loginpd) 
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

        private void Login_Load(object sender, EventArgs e)
        {
           ;
        }
    }
}
