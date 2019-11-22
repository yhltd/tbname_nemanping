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
namespace PurchasingProcedures
{
    public partial class InputCaiDanNo : Form
    {
        private Form fma;
        protected GongNeng2 gn2;
        protected string key;
        public InputCaiDanNo(Form fm,string typekey)
        {
            InitializeComponent();
            gn2 = new GongNeng2();
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            key = typekey;
            fma = fm;
        }
  
        private void button1_Click(object sender, EventArgs e)
        {
            if (!txt_caidan.Equals(string.Empty))
            {
                if (key.Equals("生成表格"))
                {
                    shengchengBiaoge scb = new shengchengBiaoge(txt_caidan.Text);
                    scb.MdiParent = fma;
                    scb.Show();
                    this.Close();
                }
                else 
                {
                    mflDgd mfl = new mflDgd(txt_caidan.Text);
                    mfl.MdiParent = fma;
                    mfl.Show();
                    this.Close();
                }
            }
            else 
            {
                MessageBox.Show("请不要输入空值！");
            }
        }

        private void InputCaiDanNo_Load(object sender, EventArgs e)
        {
            List<clsBuiness.CaiDan> cdlist = gn2.selectCaiDan("").GroupBy(c => c.CaiDanHao).Select(sc => sc.First()).ToList<clsBuiness.CaiDan>();
            txt_caidan.DataSource = cdlist;
            txt_caidan.DisplayMember = "CaiDanHao";
            txt_caidan.ValueMember = "Id";
        }
    }
}
