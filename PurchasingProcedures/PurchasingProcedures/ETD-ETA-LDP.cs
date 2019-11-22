using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PurchasingProcedures
{
    public partial class ETD_ETA_LDP : Form
    {
        public ETD_ETA_LDP()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void ETD_ETA_LDP_Load(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            //dataGridView1.Rows[1].HeaderCell.Value = "11";
            dt.Columns.Add("111", typeof(string));
            dt.Columns.Add("22", typeof(string));
            dt.Columns.Add("33", typeof(string));
            dt.Columns.Add("44", typeof(string));
            //dataGridView1.DataSource = dt;
        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
        }
    }
}
