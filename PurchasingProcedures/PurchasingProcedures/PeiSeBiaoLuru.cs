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
                    List<PeiSe> list = cal.selectPeise(cb_MianLiao.Text);
                    DataTable dt = new DataTable();
                    dateTimePicker1.Value = Convert.ToDateTime(list[0].Date);
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                            dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                    }
                    foreach (PeiSe s in list)
                    {
                        dt.Rows.Add( s.PingMing, s.HuoHao, s.GuiGe, s.C61601C1, s.C61602C1, s.C61603C1, s.C61605C1, s.C61606C1, s.C61607C1, s.C61609C1, s.C61611C1, s.C61618C1, s.C61624C1, s.C61627C1, s.C61631C1, s.C61632C1, s.C61633C1, s.C61634C1, s.MianLiaoYanSe,s.Id,s.Fabrics,s.Date);
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
            try
            {
                List<PeiSe> list = cal.selectPeise("").GroupBy(p => new { p.Fabrics }).Select(pc =>pc.First()).ToList();
                cb_MianLiao.DisplayMember = "Fabrics";
                cb_MianLiao.ValueMember = "Id";
                cb_MianLiao.DataSource = list;
            }
            catch (Exception ex) 
            {
                throw ex;
            }
            
            
        }

        private void cb_MianLiao_TextChanged(object sender, EventArgs e)
        {
       
            
        }

        private void toolStripLabel2_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = dataGridView1.DataSource as DataTable;
                cal.insertPeise(dt, cb_MianLiao.Text, dateTimePicker1.Text);
                MessageBox.Show("提交成功！");
                comboBox1_SelectedIndexChanged(sender, e);
            }
            catch (Exception ex) 
            {
                throw ex;
            }
            
        }

        private void toolStripLabel3_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult queren = MessageBox.Show("读取的'EXCEL文件'后缀必须为.Xlsx，否则读取失败！","系统提示！",MessageBoxButtons.YesNo);
                if (queren == DialogResult.Yes)
                {
                    if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        string path = openFileDialog1.FileName;
                        if (!path.Equals(string.Empty))
                        {
                            if (path.Trim().Contains("xlsx"))
                            {
                                List<PeiSe> list = cal.ReaderPeiSe(path);
                                DataTable dt = new DataTable();
                                if (list[0].Date != null)
                                {
                                    string strDate = DateTime.FromOADate(Convert.ToInt32(list[0].Date)).ToString("d");
                                    strDate = DateTime.Parse(strDate).ToString("dd-MMM-yyyy");
                                    dateTimePicker1.Value = Convert.ToDateTime(strDate);
                                }
                             
                                cb_MianLiao.DropDownStyle = ComboBoxStyle.DropDown;
                                cb_MianLiao.Text = list[0].Fabrics;
                                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                                {
                                    dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                                }
                                foreach (PeiSe s in list)
                                {
                                    if (s.PingMing != null)
                                    {
                                        dt.Rows.Add(s.PingMing, s.HuoHao, s.GuiGe, s.C61601C1, s.C61602C1, s.C61603C1, s.C61605C1, s.C61606C1, s.C61607C1, s.C61609C1, s.C61611C1, s.C61618C1, s.C61624C1, s.C61627C1, s.C61631C1, s.C61632C1, s.C61633C1, s.C61634C1, s.MianLiaoYanSe, s.Id, s.Fabrics, s.Date);
                                    }
                                }
                                dataGridView1.DataSource = dt;
                                MessageBox.Show("读取完成！");
                            }
                            else
                            {
                                MessageBox.Show("读取失败！原因:读取文件后缀非'xlsx'");
                            }
                            
                        }
                       
                        
                    }
                    
                }
            }
            catch (Exception ex) 
            {
                throw ex;

            }
            
        }
    }
}
