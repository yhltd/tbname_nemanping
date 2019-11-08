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
    public partial class DanHaoFrm : Form
    {
        protected clsAllnewLogic cal = new clsAllnewLogic();

        public DanHaoFrm()
        {
            InitializeComponent();
        }

        private void toolStripLabel1_Click(object sender, EventArgs e)
        {
            try 
            {
                List<DanHao> list = cal.SelectDanHao();
                DataTable dt = new DataTable();
                dt.Columns.Add("id", typeof(int));
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                    {
                        dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                    }
                }
                foreach (DanHao s in list)
                {
                    dt.Rows.Add(s.Id, s.CaiDanNo, s.Style, s.FABRIC_CONTENT, s.DATE, s.JiaGongChang, s.XuHao, s.Name, s.HuoHao, s.GuiGe, s.Yanse, s.Danjia, s.DanHao1, s.Jine, s.BeiZhu, s.ChangShang, s.Type);
                }
                dataGridView1.DataSource = dt;
                MessageBox.Show("刷新成功！");
            }
            catch (Exception ex) 
            {
                throw ex;
            }
        }
    }
}
