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
    public partial class Factoryinput : Form
    {
        //protected DataTable dt;
        protected List<JiaGongChang> list1;
        protected Definefactoryinput cal1;
        public Factoryinput()
        {
            InitializeComponent();
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            cal1 = new Definefactoryinput();
            list1 = new List<JiaGongChang>();
        }

        #region 刷新按钮
        private void toolStripLabel5_Click(object sender, EventArgs e)
        {
            try
            {
                list1 = cal1.selectJiaGongChang();
                DataTable dt = new DataTable();
                dt.Columns.Add("id1", typeof(int));
                dt.Columns.Add("Name1", typeof(String));
                dt.Columns.Add("Address", typeof(String));
                dt.Columns.Add("Lianxiren", typeof(String));
                dt.Columns.Add("Phone", typeof(String));
                dt.Columns.Add("ZengZhiShui", typeof(String));
                dt.Columns.Add("Kaihuhang", typeof(String));
                dt.Columns.Add("Zhanghao", typeof(String));
                foreach (JiaGongChang s in list1)
                {
                    dt.Rows.Add(s.ID, s.Name, s.Address,s.Lianxiren,s.Phone,s.ZengZhiShui,s.Kaihuhang,s.Zhanghao );
                }
                dataGridView1.DataSource = dt;
                MessageBox.Show("刷新成功");
                

            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        #endregion

       
                    #region 提交修改按钮
        private void toolStripLabel2_Click_1(object sender, EventArgs e)
        {
                DataTable dt = dataGridView1.DataSource as DataTable;
                cal1.insertJiaGongChang(dt);
            
            MessageBox.Show("提交成功！");
            toolStripLabel5_Click(sender, e);
        
        #endregion
        }

        private void toolStripLabel1_Click(object sender, EventArgs e)
        {
            DataTable dt = dataGridView1.DataSource as DataTable;
            cal1.insertJiaGongChang(dt);

            MessageBox.Show("提交成功！");
            toolStripLabel5_Click(sender, e);
        

        }

        private void toolStripLabel4_Click(object sender, EventArgs e)
        {
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string path = openFileDialog1.FileName;
                if (!path.Equals(string.Empty))
                {
                    list1 = cal1.readerJiaGongChangExcel(path);
                    DataTable dt = new DataTable();
                    dt.Columns.Add("id1", typeof(int));
                    dt.Columns.Add("Name1", typeof(String));
                    dt.Columns.Add("Address", typeof(String));
                    dt.Columns.Add("Lianxiren", typeof(String));
                    dt.Columns.Add("Phone", typeof(String));
                    dt.Columns.Add("ZengZhiShui", typeof(String));
                    dt.Columns.Add("Kaihuhang", typeof(String));
                    dt.Columns.Add("Zhanghao", typeof(String));
                    foreach (JiaGongChang s in list1)
                    {
                        dt.Rows.Add(s.ID, s.Name, s.Address, s.Lianxiren, s.Phone, s.ZengZhiShui, s.Kaihuhang, s.Zhanghao);
                    }
                    dataGridView1.DataSource = dt;
                }
                MessageBox.Show("读取成功！");

            }
        }
    }
}