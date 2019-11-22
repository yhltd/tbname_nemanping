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
        public InputCreatYjcb(Form frm)
        {
            InitializeComponent();
            cal = new clsAllnewLogic();
            df = new Definefactoryinput();
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            ETD_ETA_LDP EEL = new ETD_ETA_LDP();
            EEL.Show();
        }

        private void InputCreatYjcb_Load(object sender, EventArgs e)
        {
            List<JiaGongChang> jgc = df.selectJiaGongChang().GroupBy(j => j.Name).Select(sc => sc.First()).ToList<JiaGongChang>();
            comboBox1.DataSource = jgc;
            comboBox1.DisplayMember = "Name";
            comboBox1.ValueMember = "id";
        }
    }
}
