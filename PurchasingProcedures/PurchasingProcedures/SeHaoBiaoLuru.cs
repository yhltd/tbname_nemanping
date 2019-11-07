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
    public partial class SeHaoBiaoLuru : Form
    {
        protected clsAllnewLogic cal ;
        public SeHaoBiaoLuru()
        {
            InitializeComponent();
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            cal = new clsAllnewLogic();
        }

        private void SeHaoBiaoLuru_Load(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            List<Sehao> list = cal.selectSehao();
            dataGridView1.DataSource = list;
        }
    }
}
