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
    public partial class PeiSeBiaoLuru : Form
    {
        protected clsAllnewLogic cal = new clsAllnewLogic();
        public PeiSeBiaoLuru()
        {
            InitializeComponent();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (cb_MianLiao.SelectedValue != "")
                {
                    List<PeiSe> list = cal.selectPeise(Convert.ToInt32(cb_MianLiao.SelectedValue.ToString()));
                    DataTable dt = new DataTable();
                    dt.Columns.Add("Id", typeof(int));
                    dateTimePicker1.Value = Convert.ToDateTime(list[0].Date);
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("Id"))
                        {
                            dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                        }
                    }
                    foreach (PeiSe s in list)
                    {
                        dt.Rows.Add(s.Id, s.PingMing, s.HuoHao, s.GuiGe, s.C61601C1, s.C61602C1, s.C61603C1, s.C61605C1, s.C61606C1, s.C61607C1, s.C61609C1, s.C61611C1, s.C61618C1, s.C61624C1, s.C61627C1, s.C61631C1, s.C61632C1, s.C61633C1, s.C61634C1, s.MianLiaoYanSe);
                    }
                    dataGridView1.DataSource = dt;
                }
            }
            catch (Exception ex) 
            {
                throw ex;
            }
            
        }

        private void PeiSeBiaoLuru_Load(object sender, EventArgs e)
        {
            List<PeiSe> list = cal.selectPeise(0);
            cb_MianLiao.DisplayMember = "Fabrics";
            cb_MianLiao.ValueMember = "Id";
            cb_MianLiao.DataSource = list;
            
        }

        private void cb_MianLiao_TextChanged(object sender, EventArgs e)
        {
       
            
        }

        private void toolStripLabel2_Click(object sender, EventArgs e)
        {
            
            DataTable dt = dataGridView1.DataSource as DataTable;
            dt.Columns.Add("Fabrics", typeof(string));
            dt.Columns.Add(cb_MianLiao.Text);
            dt.Columns.Add("Id", typeof(string));
            dt.Columns.Add(cb_MianLiao.SelectedValue.ToString());
            dt.Columns.Add("Date", typeof(string));
            dt.Columns.Add(dateTimePicker1.Text);
            cal.insertPeise(dt);
        }
    }
}
