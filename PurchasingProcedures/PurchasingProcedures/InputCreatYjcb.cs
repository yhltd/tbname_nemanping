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
   
    public partial class InputCreatYjcb : Form
    {
        protected clsAllnewLogic cal;
        protected Definefactoryinput df;
        protected Form fm;
        private string cdno;
        private List<HeSuan> hesuan;
        public InputCreatYjcb(Form frm,string cd,List<HeSuan> hs)
        {
            InitializeComponent();
            cdno = cd;
            hesuan = hs;
            fm = frm;
            cal = new clsAllnewLogic();
            df = new Definefactoryinput();
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;

        }
        //private bool HaveOpened(Form _monthForm, string _childrenFormName)
        //{
        //    //查看窗口是否已经被打开
        //    bool bReturn = false;
        //    for (int i = 0; i < _monthForm.MdiChildren.Length; i++)
        //    {
        //        if (_monthForm.MdiChildren[i].Name == _childrenFormName)
        //        {
        //            _monthForm.MdiChildren[i].BringToFront();//将控件带到 Z 顺序的前面。
        //            bReturn = true;
        //            break;
        //        }
        //    }
        //    return bReturn;
        //}

        private void button1_Click(object sender, EventArgs e)
        {
            ETD_ETA_LDP EEL = new ETD_ETA_LDP(cdno, comboBox1.Text, textBox1.Text, hesuan);
            if (!EEL.IsDisposed)
            {
                //if (!HaveOpened(fm, EEL.Name))
                //{
                    //EEL.MdiParent = fm;
                    EEL.Show();
                //}
                //else 
                //{
                //    EEL.TopMost = true;
                //}
            }
        }

        private void InputCreatYjcb_Load(object sender, EventArgs e)
        {
            List<JiaGongChang> jgc = df.selectJiaGongChang().GroupBy(j => j.Name).Select(sc => sc.First()).ToList<JiaGongChang>();
            comboBox1.DataSource = jgc;
            comboBox1.DisplayMember = "Name";
            comboBox1.ValueMember = "id";
            List<clsBuiness.KuanShiBiao> cdlist = cal.SelectKuanshi().GroupBy(g => g.STYLE).Select(pc => pc.First()).ToList<clsBuiness.KuanShiBiao>();
            textBox1.DataSource = cdlist;
            textBox1.DisplayMember = "STYLE";
            textBox1.ValueMember = "Id";
        }
    }
}
