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
        //protected DataTable dt;
        protected List<Sehao> list;
        protected clsAllnewLogic cal ;
        public SeHaoBiaoLuru()
        {
            InitializeComponent();
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            cal = new clsAllnewLogic();
            list = new List<Sehao>();
        }



        #region 提交修改按钮
        private void toolStripLabel2_Click_1(object sender, EventArgs e)
        {
                DataTable dt = dataGridView1.DataSource as DataTable;
                cal.insertSehao(dt);
            
            MessageBox.Show("提交成功！");
            toolStripLabel5_Click(sender, e);
        }
        #endregion
        #region 刷新按钮
        private void toolStripLabel5_Click(object sender, EventArgs e)
        {
            try
            {
                list = cal.selectSehao();
                DataTable dt = new DataTable();
                dt.Columns.Add("Id", typeof(int));
                dt.Columns.Add("Name", typeof(String));
                dt.Columns.Add("SeHao1", typeof(String));
                foreach (Sehao s in list) 
                {
                    dt.Rows.Add(s.Id, s.Name, s.SeHao1);
                }
               dataGridView1.DataSource = dt;
               MessageBox.Show("刷新成功");
               // foreach (Sehao S in select) 
               // {
               //     Sehao sc = new Sehao();
               //     sc.Name = S.Name;
               //     sc.SeHao1 = S.SeHao1;
               //     sc.Id = S.Id;
               //     list.Add(sc);
               // }
                //Sehao sc = new Sehao();
                //list.Add(sc);
                //DataTable dt = list as DataTable();
                //dataGridView1.AllowUserToAddRows = true;
                
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        #endregion

        private void toolStripLabel4_Click(object sender, EventArgs e)
        {
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
               string path = openFileDialog1.FileName;
               if (!path.Equals(string.Empty)) 
               {
                   list = cal.readerSehaoExcel(path);
                   DataTable dt = new DataTable();
                   dt.Columns.Add("Id", typeof(int));
                   dt.Columns.Add("Name", typeof(String));
                   dt.Columns.Add("SeHao1", typeof(String));
                   foreach (Sehao s in list)
                   {
                       dt.Rows.Add(s.Id, s.Name, s.SeHao1);
                   }
                   dataGridView1.DataSource = dt;
               }
               MessageBox.Show("读取成功！");
               
            }
        }


        

   
    }
}
