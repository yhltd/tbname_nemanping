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
    public partial class Kucun : Form
    {
        protected clsAllnewLogic cal;
        public Kucun()
        {
            InitializeComponent();
            cal = new clsAllnewLogic();
        }

        private void toolStripLabel1_Click(object sender, EventArgs e)
        {
            List<KuCun> list = cal.SelectKC();
            DataTable dt = new DataTable();
            dt.Columns.Add("Id", typeof(int));
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("Id"))
                {
                    dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                }
            }
            foreach (KuCun s in list)
            {
                dt.Rows.Add(s.Id,s.PingMing,s.HuoHao,s.SeHao,s.ShuLiang,s.GongHuoFang,s.CunFangDI);
            }
            dataGridView1.DataSource = dt;
            MessageBox.Show("刷新成功！");
        }

        private void toolStripLabel2_Click(object sender, EventArgs e)
        {
            //cal.insertKucun();
        }
    }
}
