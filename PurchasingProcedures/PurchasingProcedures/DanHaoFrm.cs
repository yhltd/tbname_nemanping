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
          
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (comboBox1.SelectedValue != "")
                {
                    List<DanHao> list = cal.SelectDanHao(comboBox1.Text.ToString());
                    DataTable dt1 = new DataTable();
                    dt1.Columns.Add("Id", typeof(int));
                    DataTable dt2 = new DataTable();
                    dt2.Columns.Add("Id", typeof(int));
                    txt_STYLE.Text = list[0].Style;
                    dateTimePicker1.Text = list[0].DATE;
                    txt_mfcf.Text = list[0].FABRIC_CONTENT;
                    txt_JGC.Text = list[0].JiaGongChang;
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("Id"))
                        {
                            dt1.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                            dt2.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                        }
                    }
                    foreach (DanHao s in list)
                    {
                        if (s.Type != null)
                        {
                            if (s.Type.Equals("面料"))
                            {
                                dt1.Rows.Add(s.Id, s.CaiDanNo, s.Style, s.FABRIC_CONTENT, s.DATE, s.JiaGongChang, s.Name, s.HuoHao, s.GuiGe, s.Yanse, s.Danjia, s.DanHao1, s.Jine, s.BeiZhu, s.ChangShang, s.Type);
                            }
                            else
                            {
                                dt2.Rows.Add(s.Id, s.CaiDanNo, s.Style, s.FABRIC_CONTENT, s.DATE, s.JiaGongChang, s.Name, s.HuoHao, s.GuiGe, s.Yanse, s.Danjia, s.DanHao1, s.Jine, s.BeiZhu, s.ChangShang, s.Type);
                            }
                        }
                    }
                    dataGridView1.DataSource = dt1;
                    dataGridView2.DataSource = dt2;
                    //MessageBox.Show("刷新成功！");
                }
                
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void DanHaoFrm_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndexChanged -= comboBox1_SelectedIndexChanged;
            //List<DanHao> list = cal.SelectDanHao("").GroupBy(d =>d.CaiDanNo).Select(p => p.First).ToList<DanHao>;
            List<DanHao> list = cal.SelectDanHao("").GroupBy(d => new {d.CaiDanNo}).Select(s =>s.First()).ToList<DanHao>();
            comboBox1.DisplayMember = "CaiDanNo";
            comboBox1.ValueMember = "Id";
            comboBox1.DataSource = list;
            comboBox1.SelectedIndex = 0;//设置默认值
            comboBox1.SelectedIndexChanged += comboBox1_SelectedIndexChanged;//注册事件
            

        }

        private void toolStripLabel2_Click(object sender, EventArgs e)
        {
            //try 
            //{
                DialogResult queren = MessageBox.Show("确认提交？", "系统提示", MessageBoxButtons.YesNo);
                if (queren == DialogResult.Yes)
                {
                    DataTable DataTable1 = dataGridView1.DataSource as DataTable;
                    DataTable DataTable2 = dataGridView2.DataSource as DataTable;
                    DataTable newDataTable = DataTable1.Clone();

                    object[] obj = new object[newDataTable.Columns.Count];
                    for (int i = 0; i < DataTable1.Rows.Count; i++)
                    {
                        newDataTable.Rows.Add(DataTable1.Rows[i][0], DataTable1.Rows[i][1], DataTable1.Rows[i][2], DataTable1.Rows[i][3], DataTable1.Rows[i][4], DataTable1.Rows[i][5], DataTable1.Rows[i][6], DataTable1.Rows[i][7], DataTable1.Rows[i][8], DataTable1.Rows[i][9], DataTable1.Rows[i][10], DataTable1.Rows[i][11], DataTable1.Rows[i][12], DataTable1.Rows[i][13], DataTable1.Rows[i][14], DataTable1.Rows[i][15], "面料");
                    }

                    for (int i = 0; i < DataTable2.Rows.Count; i++)
                    {
                        newDataTable.Rows.Add(DataTable2.Rows[i][0], DataTable2.Rows[i][1], DataTable2.Rows[i][2], DataTable2.Rows[i][3], DataTable2.Rows[i][4], DataTable2.Rows[i][5], DataTable2.Rows[i][6], DataTable2.Rows[i][7], DataTable2.Rows[i][8], DataTable2.Rows[i][9], DataTable2.Rows[i][10], DataTable2.Rows[i][11], DataTable2.Rows[i][12], DataTable2.Rows[i][13], DataTable2.Rows[i][14], DataTable2.Rows[i][15], "辅料");
                    }
                    foreach (DataRow dr in newDataTable.Rows)
                    {
                        if (dr[1] is DBNull || dr[1].Equals(string.Empty))
                        {
                            dr[1] = comboBox1.Text;
                        }
                        if (dr[2] is DBNull || dr[2].Equals(string.Empty))
                        {
                            dr[2] = txt_STYLE.Text;
                        }
                        if (dr[3] is DBNull || dr[3].Equals(string.Empty))
                        {
                            dr[3] = txt_mfcf.Text;
                        }
                        if (dr[4] is DBNull || dr[4].Equals(string.Empty))
                        {
                            dr[4] = dateTimePicker1.Text;
                        }
                        if (dr[5] is DBNull || dr[5].Equals(string.Empty))
                        {
                            dr[5] = txt_JGC.Text;
                        }
                    }
                    cal.insertDanhao(newDataTable);
                    MessageBox.Show("提交成功！");
                }
            //}
            //catch (Exception ex) 
            //{
            //    throw ex;
            //}
           
        }

        private void toolStripLabel3_Click(object sender, EventArgs e)
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
                                List<DanHao> list = cal.Readerdh(path);
                                DataTable dt1 = new DataTable();
                                dt1.Columns.Add("Id", typeof(int));
                                DataTable dt2 = new DataTable();
                                dt2.Columns.Add("Id", typeof(int));
                                txt_STYLE.Text = list[0].Style;
                                dateTimePicker1.Text = list[0].DATE;
                                txt_mfcf.Text = list[0].FABRIC_CONTENT;
                                txt_JGC.Text = list[0].JiaGongChang;
                                this.comboBox1.DropDownStyle = ComboBoxStyle.DropDown;
                                this.comboBox1.Text = list[0].CaiDanNo;
                                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                                {
                                    if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("Id"))
                                    {
                                        dt1.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                                        dt2.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                                    }
                                }
                                foreach (DanHao s in list)
                                {
                                    if (s.Type != null)
                                    {
                                        if (!s.Name.Contains("金额") && !s.Name.Equals("面料") && !s.Name.Contains("辅料") && !s.Name.Equals(string.Empty)) 
                                        {
                                            if (s.Type.Equals("面料"))
                                            {
                                                dt1.Rows.Add(s.Id, s.CaiDanNo, s.Style, s.FABRIC_CONTENT, s.DATE, s.JiaGongChang, s.Name, s.HuoHao, s.GuiGe, s.Yanse, s.Danjia, s.DanHao1, s.Jine, s.BeiZhu, s.ChangShang, s.Type);
                                            }
                                            else
                                            {
                                                dt2.Rows.Add(s.Id, s.CaiDanNo, s.Style, s.FABRIC_CONTENT, s.DATE, s.JiaGongChang, s.Name, s.HuoHao, s.GuiGe, s.Yanse, s.Danjia, s.DanHao1, s.Jine, s.BeiZhu, s.ChangShang, s.Type);
                                            }

                                        }
                                    }
                                }
                                dataGridView1.DataSource = dt1;
                                dataGridView2.DataSource = dt2;
                            }
                            else
                            {
                                MessageBox.Show("读取失败！原因:读取文件后缀非'xlsx'");
                            }
                        }
                    }
                }
           
        }
    }
}
