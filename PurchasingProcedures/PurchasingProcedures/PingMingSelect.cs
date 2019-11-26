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

namespace PurchasingProcedures
{
    public partial class PingMingSelect : Form
    {
        protected clsAllnewLogic cal;
        protected mflDgd f;
        protected string cdhao;
        protected GongNeng2 gn;
        protected Define1 df;
        protected string jgc;
        protected string ks;
        public PingMingSelect(mflDgd fm,string cd )
        {
            cal = new clsAllnewLogic();
            gn = new GongNeng2();
            df = new Define1();
            //jgc = jiagongchang;
            //ks = kuanshi;
            cdhao = cd;
            f = fm;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;

            InitializeComponent();
        }

        private void PingMingSelect_Load(object sender, EventArgs e)
        {
            //List<clsBuiness.DanHao> dh = cal.SelectDanHao("");
            List<clsBuiness.DanHao> list = cal.SelectDanHao("").FindAll(d => d.CaiDanNo.Trim().Equals(cdhao)).GroupBy(gp => gp.Name.Trim()).Select(s => s.First()).ToList<DanHao>();

            comboBox1.DataSource = list;
            comboBox1.DisplayMember = "Name";
            comboBox1.ValueMember = "Id";
        }
        private void button1_Click(object sender, EventArgs e)
        {

            f.ChuanHuiMFL = cal.SelectMianFuLiao().FindAll(fc=> fc.PingMing.Equals(comboBox1.Text));
            //f.pinming = comboBox1.Text;
            //f.hesuan = CreateFuLiao(this.comboBox1.Text, "辅料");
            if (f.ChuanHuiMFL.Count > 0)
            {
                f.mflDgd_Load(sender, e);
                f.Visible = true;
            }
            else 
            {
                MessageBox.Show("查询失败！原因：该品名内 无 信息 ");
            }
            //this.Close();
        }
    }
}
