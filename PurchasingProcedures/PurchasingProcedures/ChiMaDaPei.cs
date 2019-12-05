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
using System.Threading;
using System.Runtime.InteropServices;
using System.IO;
namespace PurchasingProcedures
{
    public partial class ChiMaDaPei : Form
    {
        [DllImport("kernel32.dll")]
        public static extern IntPtr _lopen(string lpPathName, int iReadWrite);
        public const int OF_READWRITE = 2;
        public const int OF_SHARE_DENY_NONE = 0x40;
        public readonly IntPtr HFILE_ERROR = new IntPtr(-1);
        public clsAllnewLogic cal = new clsAllnewLogic();
        [DllImport("kernel32.dll")]
        public static extern bool CloseHandle(IntPtr hObject);

        public ChiMaDaPei()
        {
            InitializeComponent();
            //this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
        }
        public List<ChiMa_Dapeibiao> bindDataGirdview(string wheres) 
        {
            List<ChiMa_Dapeibiao> list = new List<ChiMa_Dapeibiao>();
            if (wheres.Equals(string.Empty))
            {
               list = cal.SelectChiMaDapei("");
            }
            else 
            {
                list = cal.SelectChiMaDapei(wheres);
            }
            return list;
            
        }
        private void toolStripLabel1_Click(object sender, EventArgs e)
        {
            
        }

        private void toolStripLabel2_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult dr = MessageBox.Show("确认要提交吗？","信息",MessageBoxButtons.YesNo);
                if (dr == DialogResult.Yes)
                {
                    this.backgroundWorker1.RunWorkerAsync(); // 运行 backgroundWorker 组件

                    JingDu form = new JingDu(this.backgroundWorker1, "提交中");// 显示进度条窗体
                    form.ShowDialog(this);
                    form.Close();

                    DataTable dt = dataGridView1.DataSource as DataTable;
                    #region RGL1
                    
                        dt = new DataTable();
                        //dt.Columns.Add("id", typeof(int));
                        for (int i = 0; i < dataGridView1.Columns.Count; i++)
                        {
                            if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                            {
                                dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                            }
                        }
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            if (dataGridView1.Rows[i].Cells[6].Value != null)
                            {
                                dt.Rows.Add( dataGridView1.Rows[i].Cells[0].Value, dataGridView1.Rows[i].Cells[1].Value, dataGridView1.Rows[i].Cells[2].Value, dataGridView1.Rows[i].Cells[3].Value, dataGridView1.Rows[i].Cells[4].Value, dataGridView1.Rows[i].Cells[5].Value, dataGridView1.Rows[i].Cells[6].Value, dataGridView1.Rows[i].Cells[7].Value, dataGridView1.Rows[i].Cells[8].Value, dataGridView1.Rows[i].Cells[9].Value, dataGridView1.Rows[i].Cells[10].Value, dataGridView1.Rows[i].Cells[11].Value, dataGridView1.Rows[i].Cells[12].Value, dataGridView1.Rows[i].Cells[13].Value, dataGridView1.Rows[i].Cells[14].Value, dataGridView1.Rows[i].Cells[15].Value, dataGridView1.Rows[i].Cells[16].Value, dataGridView1.Rows[i].Cells[17].Value, dataGridView1.Rows[i].Cells[18].Value, dataGridView1.Rows[i].Cells[19].Value, dataGridView1.Rows[i].Cells[20].Value, dataGridView1.Rows[i].Cells[21].Value, dataGridView1.Rows[i].Cells[22].Value, dataGridView1.Rows[i].Cells[23].Value, dataGridView1.Rows[i].Cells[24].Value, dataGridView1.Rows[i].Cells[25].Value, dataGridView1.Rows[i].Cells[26].Value, dataGridView1.Rows[i].Cells[27].Value, dataGridView1.Rows[i].Cells[28].Value, dataGridView1.Rows[i].Cells[29].Value, dataGridView1.Rows[i].Cells[30].Value, dataGridView1.Rows[i].Cells[31].Value, dataGridView1.Rows[i].Cells[32].Value, dataGridView1.Rows[i].Cells[33].Value, dataGridView1.Rows[i].Cells[34].Value, dataGridView1.Rows[i].Cells[35].Value, dataGridView1.Rows[i].Cells[36].Value, dataGridView1.Rows[i].Cells[37].Value, dataGridView1.Rows[i].Cells[38].Value, dataGridView1.Rows[i].Cells[39].Value, dataGridView1.Rows[i].Cells[40].Value, dataGridView1.Rows[i].Cells[41].Value, dataGridView1.Rows[i].Cells[42].Value);
                            }
                        }
                    
                    cal.InsertChima(dt, "RGL1");
                    #endregion

                    #region RGL2
                    dt = new DataTable();
                    //dt.Columns.Add("id", typeof(int));
                    for (int i = 0; i < RGL2.Columns.Count; i++)
                    {
                        if (!RGL2.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                        {
                            dt.Columns.Add(RGL2.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                        }
                    }
                    for (int i = 0; i < RGL2.Rows.Count; i++)
                    {
                        dt.Rows.Add(RGL2.Rows[i].Cells[0].Value, RGL2.Rows[i].Cells[1].Value, RGL2.Rows[i].Cells[2].Value, RGL2.Rows[i].Cells[3].Value, RGL2.Rows[i].Cells[4].Value, RGL2.Rows[i].Cells[5].Value, RGL2.Rows[i].Cells[6].Value, RGL2.Rows[i].Cells[7].Value, RGL2.Rows[i].Cells[8].Value, RGL2.Rows[i].Cells[9].Value, RGL2.Rows[i].Cells[10].Value, RGL2.Rows[i].Cells[11].Value, RGL2.Rows[i].Cells[12].Value, RGL2.Rows[i].Cells[13].Value, RGL2.Rows[i].Cells[14].Value, RGL2.Rows[i].Cells[15].Value, RGL2.Rows[i].Cells[16].Value, RGL2.Rows[i].Cells[17].Value, RGL2.Rows[i].Cells[18].Value, RGL2.Rows[i].Cells[19].Value, RGL2.Rows[i].Cells[20].Value, RGL2.Rows[i].Cells[21].Value, RGL2.Rows[i].Cells[22].Value, RGL2.Rows[i].Cells[23].Value, RGL2.Rows[i].Cells[24].Value, RGL2.Rows[i].Cells[25].Value, RGL2.Rows[i].Cells[26].Value, RGL2.Rows[i].Cells[27].Value, RGL2.Rows[i].Cells[28].Value, RGL2.Rows[i].Cells[29].Value, RGL2.Rows[i].Cells[30].Value, RGL2.Rows[i].Cells[31].Value, RGL2.Rows[i].Cells[32].Value, RGL2.Rows[i].Cells[33].Value, RGL2.Rows[i].Cells[34].Value, RGL2.Rows[i].Cells[35].Value, RGL2.Rows[i].Cells[36].Value, RGL2.Rows[i].Cells[37].Value, RGL2.Rows[i].Cells[38].Value, RGL2.Rows[i].Cells[39].Value, RGL2.Rows[i].Cells[40].Value, RGL2.Rows[i].Cells[41].Value, RGL2.Rows[i].Cells[42].Value);
                    }
                    //MessageBox.Show("提交成功！");
                    cal.InsertChima2(dt, "RGL2");
                    #endregion

                    #region SLIM
                    dt = new DataTable();
                    //dt.Columns.Add("id", typeof(int));
                    for (int i = 0; i < SLIM.Columns.Count; i++)
                    {
                        if (!SLIM.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                        {
                            dt.Columns.Add(SLIM.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                        }
                    }
                    for (int i = 0; i < SLIM.Rows.Count; i++)
                    {
                        dt.Rows.Add(SLIM.Rows[i].Cells[0].Value, SLIM.Rows[i].Cells[1].Value, SLIM.Rows[i].Cells[2].Value, SLIM.Rows[i].Cells[3].Value, SLIM.Rows[i].Cells[4].Value, SLIM.Rows[i].Cells[5].Value, SLIM.Rows[i].Cells[6].Value, SLIM.Rows[i].Cells[7].Value, SLIM.Rows[i].Cells[8].Value, SLIM.Rows[i].Cells[9].Value, SLIM.Rows[i].Cells[10].Value, SLIM.Rows[i].Cells[11].Value, SLIM.Rows[i].Cells[12].Value, SLIM.Rows[i].Cells[13].Value, SLIM.Rows[i].Cells[14].Value, SLIM.Rows[i].Cells[15].Value, SLIM.Rows[i].Cells[16].Value, SLIM.Rows[i].Cells[17].Value, SLIM.Rows[i].Cells[18].Value, SLIM.Rows[i].Cells[19].Value, SLIM.Rows[i].Cells[20].Value, SLIM.Rows[i].Cells[21].Value, SLIM.Rows[i].Cells[22].Value, SLIM.Rows[i].Cells[23].Value, SLIM.Rows[i].Cells[24].Value, SLIM.Rows[i].Cells[25].Value, SLIM.Rows[i].Cells[26].Value, SLIM.Rows[i].Cells[27].Value, SLIM.Rows[i].Cells[28].Value);
                    }
                    cal.InsertChima3(dt, "SLIM");
                    #endregion

                    #region RGLJ
                    dt = new DataTable();
                    //dt.Columns.Add("id", typeof(int));
                    for (int i = 0; i < RGLJ.Columns.Count; i++)
                    {
                        if (!RGLJ.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                        {
                            dt.Columns.Add(RGLJ.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                        }
                    }
                    for (int i = 0; i < RGLJ.Rows.Count; i++)
                    {
                        dt.Rows.Add(RGLJ.Rows[i].Cells[0].Value, RGLJ.Rows[i].Cells[1].Value, RGLJ.Rows[i].Cells[2].Value, RGLJ.Rows[i].Cells[3].Value, RGLJ.Rows[i].Cells[4].Value, RGLJ.Rows[i].Cells[5].Value, RGLJ.Rows[i].Cells[6].Value, RGLJ.Rows[i].Cells[7].Value, RGLJ.Rows[i].Cells[8].Value, RGLJ.Rows[i].Cells[9].Value, RGLJ.Rows[i].Cells[10].Value, RGLJ.Rows[i].Cells[11].Value, RGLJ.Rows[i].Cells[12].Value, RGLJ.Rows[i].Cells[13].Value, RGLJ.Rows[i].Cells[14].Value, RGLJ.Rows[i].Cells[15].Value, RGLJ.Rows[i].Cells[16].Value, RGLJ.Rows[i].Cells[17].Value, RGLJ.Rows[i].Cells[18].Value, RGLJ.Rows[i].Cells[19].Value, RGLJ.Rows[i].Cells[20].Value, RGLJ.Rows[i].Cells[21].Value, RGLJ.Rows[i].Cells[22].Value, RGLJ.Rows[i].Cells[23].Value, RGLJ.Rows[i].Cells[24].Value, RGLJ.Rows[i].Cells[25].Value, RGLJ.Rows[i].Cells[26].Value, RGLJ.Rows[i].Cells[27].Value, RGLJ.Rows[i].Cells[28].Value, RGLJ.Rows[i].Cells[29].Value, RGLJ.Rows[i].Cells[30].Value, RGLJ.Rows[i].Cells[31].Value, RGLJ.Rows[i].Cells[32].Value, RGLJ.Rows[i].Cells[33].Value, RGLJ.Rows[i].Cells[34].Value, RGLJ.Rows[i].Cells[35].Value, RGLJ.Rows[i].Cells[36].Value, RGLJ.Rows[i].Cells[37].Value, RGLJ.Rows[i].Cells[38].Value, RGLJ.Rows[i].Cells[39].Value, RGLJ.Rows[i].Cells[40].Value, RGLJ.Rows[i].Cells[41].Value, RGLJ.Rows[i].Cells[42].Value);
                    }
                    cal.InsertChima4(dt, "RGLJ");
                    #endregion

                    #region D.PANT
                    dt = new DataTable();
                    //dt.Columns.Add("id", typeof(int));
                    for (int i = 0; i < headerUnitView1.Columns.Count; i++)
                    {
                        if (!headerUnitView1.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                        {
                            dt.Columns.Add(headerUnitView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                        }
                    }
                    for (int i = 0; i < headerUnitView1.Rows.Count; i++)
                    {
                        dt.Rows.Add(headerUnitView1.Rows[i].Cells[0].Value, headerUnitView1.Rows[i].Cells[1].Value, headerUnitView1.Rows[i].Cells[2].Value, headerUnitView1.Rows[i].Cells[3].Value, headerUnitView1.Rows[i].Cells[4].Value, headerUnitView1.Rows[i].Cells[5].Value, headerUnitView1.Rows[i].Cells[6].Value, headerUnitView1.Rows[i].Cells[7].Value, headerUnitView1.Rows[i].Cells[8].Value, headerUnitView1.Rows[i].Cells[9].Value, headerUnitView1.Rows[i].Cells[10].Value, headerUnitView1.Rows[i].Cells[11].Value, headerUnitView1.Rows[i].Cells[12].Value, headerUnitView1.Rows[i].Cells[13].Value, headerUnitView1.Rows[i].Cells[14].Value, headerUnitView1.Rows[i].Cells[15].Value, headerUnitView1.Rows[i].Cells[16].Value, headerUnitView1.Rows[i].Cells[17].Value, headerUnitView1.Rows[i].Cells[18].Value, headerUnitView1.Rows[i].Cells[19].Value, headerUnitView1.Rows[i].Cells[20].Value, headerUnitView1.Rows[i].Cells[21].Value, headerUnitView1.Rows[i].Cells[22].Value, headerUnitView1.Rows[i].Cells[23].Value, headerUnitView1.Rows[i].Cells[24].Value, headerUnitView1.Rows[i].Cells[25].Value, headerUnitView1.Rows[i].Cells[26].Value, headerUnitView1.Rows[i].Cells[27].Value, headerUnitView1.Rows[i].Cells[28].Value, headerUnitView1.Rows[i].Cells[29].Value, headerUnitView1.Rows[i].Cells[30].Value, headerUnitView1.Rows[i].Cells[31].Value, headerUnitView1.Rows[i].Cells[32].Value, headerUnitView1.Rows[i].Cells[33].Value, headerUnitView1.Rows[i].Cells[34].Value, headerUnitView1.Rows[i].Cells[35].Value, headerUnitView1.Rows[i].Cells[36].Value, headerUnitView1.Rows[i].Cells[37].Value, headerUnitView1.Rows[i].Cells[38].Value, headerUnitView1.Rows[i].Cells[39].Value, headerUnitView1.Rows[i].Cells[40].Value, headerUnitView1.Rows[i].Cells[41].Value);
                    }
                    cal.InsertChima5(dt, "D.PANT");
                    #endregion

                    #region C.PANT
                    dt = new DataTable();
                    //dt.Columns.Add("id", typeof(int));
                    for (int i = 0; i < headerUnitView2.Columns.Count; i++)
                    {
                        if (!headerUnitView2.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                        {
                            dt.Columns.Add(headerUnitView2.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                        }
                    }
                    for (int i = 0; i < headerUnitView2.Rows.Count; i++)
                    {
                        dt.Rows.Add(headerUnitView2.Rows[i].Cells[0].Value, headerUnitView2.Rows[i].Cells[1].Value, headerUnitView2.Rows[i].Cells[2].Value, headerUnitView2.Rows[i].Cells[3].Value, headerUnitView2.Rows[i].Cells[4].Value, headerUnitView2.Rows[i].Cells[5].Value, headerUnitView2.Rows[i].Cells[6].Value, headerUnitView2.Rows[i].Cells[7].Value, headerUnitView2.Rows[i].Cells[8].Value, headerUnitView2.Rows[i].Cells[9].Value, headerUnitView2.Rows[i].Cells[10].Value, headerUnitView2.Rows[i].Cells[11].Value, headerUnitView2.Rows[i].Cells[12].Value, headerUnitView2.Rows[i].Cells[13].Value, headerUnitView2.Rows[i].Cells[14].Value, headerUnitView2.Rows[i].Cells[15].Value, headerUnitView2.Rows[i].Cells[16].Value, headerUnitView2.Rows[i].Cells[17].Value, headerUnitView2.Rows[i].Cells[18].Value, headerUnitView2.Rows[i].Cells[19].Value, headerUnitView2.Rows[i].Cells[20].Value, headerUnitView2.Rows[i].Cells[21].Value, headerUnitView2.Rows[i].Cells[22].Value, headerUnitView2.Rows[i].Cells[23].Value, headerUnitView2.Rows[i].Cells[24].Value, headerUnitView2.Rows[i].Cells[25].Value, headerUnitView2.Rows[i].Cells[26].Value, headerUnitView2.Rows[i].Cells[27].Value, headerUnitView2.Rows[i].Cells[28].Value, headerUnitView2.Rows[i].Cells[29].Value, headerUnitView2.Rows[i].Cells[30].Value, headerUnitView2.Rows[i].Cells[31].Value, headerUnitView2.Rows[i].Cells[32].Value, headerUnitView2.Rows[i].Cells[33].Value, headerUnitView2.Rows[i].Cells[34].Value, headerUnitView2.Rows[i].Cells[35].Value, headerUnitView2.Rows[i].Cells[36].Value, headerUnitView2.Rows[i].Cells[37].Value, headerUnitView2.Rows[i].Cells[38].Value, headerUnitView2.Rows[i].Cells[39].Value, headerUnitView2.Rows[i].Cells[40].Value, headerUnitView2.Rows[i].Cells[41].Value);
                    }

                    cal.InsertChima6(dt, "C.PANT");
                    #endregion

                    MessageBox.Show("提交成功！");
                    ChiMaDaPei_Load(sender, e);
                    //bindDataGirdview(comboBox1.Text);
                }
            }
            catch (Exception ex) 
            {
                //throw ex;
                MessageBox.Show(ex.Message);
            }
            
        }

        private void toolStripLabel3_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult queren = MessageBox.Show("读取的'EXCEL文件'后缀必须为.Xlsx，否则读取失败！", "系统提示！", MessageBoxButtons.YesNo);
                if (queren == DialogResult.Yes)
                {
                    List<ChiMa_Dapeibiao> list = new List<ChiMa_Dapeibiao>();
                    List<RGL2> R= new List<RGL2>();
                    List<SLIM> S = new List<SLIM>();
                    List<RGLJ> RJ = new List<RGLJ>();
                    List<D_PANT> D = new List<D_PANT>();
                    List<C_PANT> C = new List<C_PANT>();
                    if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        string path = openFileDialog1.FileName;
                        if (!path.Equals(string.Empty))
                        {
                            if (!File.Exists(path))
                            {
                                MessageBox.Show("文件不存在！");
                                return;
                            }
                            IntPtr vHandle = _lopen(path, OF_READWRITE | OF_SHARE_DENY_NONE);
                            if (vHandle == HFILE_ERROR)
                            {
                                MessageBox.Show("文件被占用！");
                                return;
                            }
                            CloseHandle(vHandle);
                            if (path.Trim().Contains("xlsx"))
                            {
                                this.backgroundWorker1.RunWorkerAsync(); // 运行 backgroundWorker 组件
                                JingDu form = new JingDu(this.backgroundWorker1, "读取中");// 显示进度条窗体
                                form.ShowDialog(this);
                                form.Close();
                                #region RGL1
                                list = cal.ReaderChiMaDapei(path);

                                DataTable dt = new DataTable();
                                //dt.Columns.Add("id", typeof(int));
                                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                                {
                                    if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                                    {
                                        dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                                    }
                                }
                                foreach (ChiMa_Dapeibiao s in list)
                                {
                                    dt.Rows.Add(s.id, s.LOT__面料, s.STYLE_款式, s.ART_货号, s.COLOR_颜色, s.COLOR__颜色编号, s.JACKET_上衣_PANT_裤子, s.C34R_28, s.C36R_30, s.C38R_32, s.C40R___34, s.C42R_36, s.C44R_38, s.C46R_40, s.C48R_42, s.C50R_44, s.C52R_46, s.C54R_48, s.C56R_50, s.C58R_52, s.C60R_54, s.C62R_56, s.C36L_30, s.C38L_32, s.C40L_34, s.C42L_36, s.C44L_38, s.C46L_40, s.C48L_42, s.C50L_44, s.C52L_46, s.C54L_48, s.C56L_50, s.C58L_52, s.C60L_54, s.C62L_56, s.C34S_28, s.C36S_30, s.C38S_32, s.C40S_34, s.C42S_36, s.C44S_38, s.C46S_40, s.DingdanHeji);
                                }
                                dataGridView1.DataSource = dt; 
                                #endregion

                                #region RGL2
                                R = cal.ReaderChiMaDapei2(path);
                                dt = new DataTable();
                                //dt.Columns.Add("id", typeof(int));
                                for (int i = 0; i < RGL2.Columns.Count; i++)
                                {
                                    if (!RGL2.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                                    {
                                        dt.Columns.Add(RGL2.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                                    }
                                }
                                foreach (RGL2 s in R)
                                {
                                    dt.Rows.Add(s.LOT_, s.STYLE_, s.ART, s.COLOR, s.COLORName, s.shangyi_kuzi, s.C34R, s.C36R, s.C38R, s.C40R, s.C42R, s.C44R, s.C46R, s.C48R, s.C50R, s.C52R, s.C54R, s.C56R, s.C58R, s.C60R, s.C62R, s.C36L, s.C38L, s.C40L, s.C42L, s.C44L, s.C46L, s.C48L, s.C50L, s.C52L, s.C54L, s.C56L, s.C58L, s.C60L, s.C62L, s.C34S, s.C36S, s.C38S, s.C40S, s.C42S, s.C44S, s.C46S, s.Sub_Total);
                                }
                                RGL2.DataSource = dt; 

                                #endregion

                                #region SLIM
                                    S = cal.ReaderChiMaDapei3(path);
                                    dt = new DataTable();
                                    //dt.Columns.Add("id", typeof(int));
                                    for (int i = 0; i < SLIM.Columns.Count; i++)
                                    {
                                        if (!SLIM.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                                        {
                                            dt.Columns.Add(SLIM.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                                        }
                                    }
                                    foreach (SLIM s in S)
                                    {
                                        dt.Rows.Add(s.LOT_, s.STYLE_, s.ART, s.COLOR, s.COLOR_, s.shangyi_kuzi, s.C34R, s.C36R, s.C38R, s.C40R, s.C42R, s.C44R, s.C46R, s.C48R, s.C36L, s.C38L, s.C40L, s.C42L, s.C44L, s.C46L, s.C48L, s.C34S, s.C36S, s.C38S, s.C40S, s.C42S, s.C44S, s.C46S, s.Sub_Total);
                                    }
                                    SLIM.DataSource = dt; 
                                #endregion

                                #region RGLJ
                                    RJ = cal.ReaderChiMaDapei4(path);
                                    dt = new DataTable();
                                    //dt.Columns.Add("id", typeof(int));
                                    for (int i = 0; i < RGLJ.Columns.Count; i++)
                                    {
                                        if (!RGLJ.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                                        {
                                            dt.Columns.Add(RGLJ.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                                        }
                                    }
                                    foreach (RGLJ s in RJ)
                                    {
                                        dt.Rows.Add(s.LOT_, s.STYLE_, s.ART, s.COLOR, s.COLOR_, s.shangyi, s.C34R, s.C36R, s.C38R, s.C40R, s.C42R, s.C44R, s.C46R, s.C48R, s.C50R, s.C52R, s.C54R, s.C56R, s.C58R, s.C60R, s.C62R, s.C36L, s.C38L, s.C40L, s.C42L, s.C44L, s.C46L, s.C48L, s.C50L, s.C52L, s.C54L, s.C56L, s.C58L, s.C60L, s.C62L, s.C34S, s.C36S, s.C38S, s.C40S, s.C42S, s.C44S, s.C46S, s.Sub_Total);
                                    }
                                    RGLJ.DataSource = dt; 

                                #endregion

                                #region D.PANT
                                    D = cal.ReaderChiMaDapei5(path);
                                    dt = new DataTable();
                                    //dt.Columns.Add("id", typeof(int));
                                    for (int i = 0; i < headerUnitView1.Columns.Count; i++)
                                    {
                                        if (!headerUnitView1.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                                        {
                                            dt.Columns.Add(headerUnitView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                                        }
                                    }
                                    int RI = 0;
                                    foreach (D_PANT s in D)
                                    {
                                        headerUnitView1.Rows.Add();
                                        headerUnitView1.Rows[0].Cells["LOT"].Value = s.LOT_;
                                        headerUnitView1.Rows[0].Cells["STYLE"].Value = s.STYLE_;
                                        headerUnitView1.Rows[0].Cells["ART"].Value = s.ART;
                                        headerUnitView1.Rows[0].Cells["COLOR"].Value = s.COLOR;
                                        headerUnitView1.Rows[0].Cells["COLORName"].Value = s.COLORName;
                                        headerUnitView1.Rows[0].Cells["yaowei"].Value = s.yaowei;
                                        headerUnitView1.Rows[0].Cells["C30W_R_30L"].Value = s.C30W_R_30L;
                                        headerUnitView1.Rows[0].Cells["C30W_L_32L"].Value = s.C30W_L_32L;
                                        headerUnitView1.Rows[0].Cells["C32W_R_30L"].Value = s.C32W_R_30L;
                                        headerUnitView1.Rows[0].Cells["C32W_L_32L"].Value = s.C32W_L_32L;
                                        headerUnitView1.Rows[0].Cells["C34W_S_38L"].Value = s.C34W_S_38L;
                                        headerUnitView1.Rows[0].Cells["C34W_S_39L"].Value = s.C34W_S_39L;
                                        headerUnitView1.Rows[0].Cells["C34W_R_30L"].Value = s.C34W_R_30L;
                                        headerUnitView1.Rows[0].Cells["C34W_L_32L"].Value = s.C34W_L_32L;
                                        headerUnitView1.Rows[0].Cells["C34W_L_34L"].Value = s.C34W_L_34L;
                                        headerUnitView1.Rows[0].Cells["C36W_S_28L"].Value = s.C36W_S_28L;
                                        headerUnitView1.Rows[0].Cells["C36W_S_29L"].Value = s.C36W_S_29L;
                                        headerUnitView1.Rows[0].Cells["C36W_R_30L"].Value = s.C36W_R_30L;
                                        headerUnitView1.Rows[0].Cells["C36W_R_31L"].Value = s.C36W_R_31L;
                                        headerUnitView1.Rows[0].Cells["C38W_S_28L"].Value = s.C38W_S_28L;
                                        headerUnitView1.Rows[0].Cells["C38W_R_30L"].Value = s.C38W_R_30L;
                                        headerUnitView1.Rows[0].Cells["C38W_R_31L"].Value = s.C38W_R_31L;
                                        headerUnitView1.Rows[0].Cells["C38W_L_32L"].Value = s.C38W_L_32L;
                                        headerUnitView1.Rows[0].Cells["C38W_L_34L"].Value = s.C38W_L_34L;
                                        headerUnitView1.Rows[0].Cells["C40W_S_28L"].Value = s.C40W_S_28L;
                                        headerUnitView1.Rows[0].Cells["C40W_S_29L"].Value = s.C40W_S_29L;
                                        headerUnitView1.Rows[0].Cells["C40W_R_30L"].Value = s.C40W_R_30L;
                                        headerUnitView1.Rows[0].Cells["C40W_R_31L"].Value = s.C40W_R_31L;
                                        headerUnitView1.Rows[0].Cells["C40W_L_32L"].Value = s.C40W_L_32L;
                                        headerUnitView1.Rows[0].Cells["C40W_L_34L"].Value = s.C40W_L_34L;
                                        headerUnitView1.Rows[0].Cells["C42W_R_30L"].Value = s.C42W_R_30L;
                                        headerUnitView1.Rows[0].Cells["C42W_L_32L"].Value = s.C42W_L_32L;
                                        headerUnitView1.Rows[0].Cells["C42W_L_34L"].Value = s.C42W_L_34L;
                                        headerUnitView1.Rows[0].Cells["C44W_R_30L"].Value = s.C44W_R_30L;
                                        headerUnitView1.Rows[0].Cells["C44W_L_32L"].Value = s.C44W_L_32L;
                                        headerUnitView1.Rows[0].Cells["C44W_L_34L"].Value = s.C44W_L_34L;
                                        headerUnitView1.Rows[0].Cells["C46W_R_30L"].Value = s.C46W_R_30L;
                                        headerUnitView1.Rows[0].Cells["C46W_L_32L"].Value = s.C46W_L_32L;
                                        headerUnitView1.Rows[0].Cells["C48W_R_30L"].Value = s.C48W_R_30L;
                                        headerUnitView1.Rows[0].Cells["C48W_L_32L"].Value = s.C48W_L_32L;
                                        headerUnitView1.Rows[0].Cells["C50W_L_32L"].Value = s.C50W_L_32L;
                                        headerUnitView1.Rows[0].Cells["Sub_Total"].Value = s.Sub_Total;
                                        //headerUnitView1.Rows[0].Cells["ART"].Value = s.ART;
                                        //dt.Rows.Add(s.LOT_, s.STYLE_, s.ART, s.COLOR, s.COLORName, s.yaowei, s.C30W_R_30L, s.C30W_L_32L, s.C32W_R_30L, s.C32W_L_32L, s.C34W_S_38L, s.C34W_S_39L, s.C34W_R_30L, s.C34W_L_32L, s.C34W_L_34L, s.C36W_S_28L, s.C36W_S_29L, s.C36W_R_30L, s.C36W_R_31L, s.C38W_S_28L, s.C38W_R_30L, s.C38W_R_31L, s.C38W_L_32L, s.C38W_L_34L, s.C40W_S_28L, s.C40W_S_29L, s.C40W_R_30L, s.C40W_R_31L, s.C40W_L_32L, s.C40W_L_34L, s.C42W_R_30L, s.C42W_L_32L, s.C42W_L_34L, s.C44W_R_30L, s.C44W_L_32L, s.C44W_L_34L, s.C46W_R_30L, s.C46W_L_32L, s.C48W_R_30L, s.C48W_L_32L, s.C50W_L_32L, s.Sub_Total);
                                    }
                                    //headerUnitView1.DataSource = dt; 
                                #endregion

                                #region C.PANT
                                    C = cal.ReaderChiMaDapei6(path);
                                    dt = new DataTable();
                                    //dt.Columns.Add("id", typeof(int));
                                    for (int i = 0; i < headerUnitView2.Columns.Count; i++)
                                    {
                                        if (!headerUnitView2.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                                        {
                                            dt.Columns.Add(headerUnitView2.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                                        }
                                    }
                                    foreach (C_PANT s in C)
                                    {
                                        headerUnitView2.Rows.Add();
                                        headerUnitView2.Rows[0].Cells["LOT_"].Value = s.LOT_;
                                        headerUnitView2.Rows[0].Cells["STYLE_"].Value = s.STYLE_;
                                        headerUnitView2.Rows[0].Cells["ART_"].Value = s.ART;
                                        headerUnitView2.Rows[0].Cells["COLOR_"].Value = s.COLOR;
                                        headerUnitView2.Rows[0].Cells["COLORName_"].Value = s.COLORName;
                                        headerUnitView2.Rows[0].Cells["yaowei_"].Value = s.yaowei;
                                        headerUnitView2.Rows[0].Cells["C30W_29L"].Value = s.C30W_29L;
                                        headerUnitView2.Rows[0].Cells["C30W_30L"].Value = s.C30W_30L;
                                        headerUnitView2.Rows[0].Cells["C30W_32L"].Value = s.C30W_32L;
                                        headerUnitView2.Rows[0].Cells["C31W_30L"].Value = s.C31W_30L;
                                        headerUnitView2.Rows[0].Cells["C31W_32L"].Value = s.C31W_32L;
                                        headerUnitView2.Rows[0].Cells["C32W_28L"].Value = s.C32W_28L;
                                        headerUnitView2.Rows[0].Cells["C32W_30L"].Value = s.C32W_30L;
                                        headerUnitView2.Rows[0].Cells["C32W_32L"].Value = s.C32W_32L;
                                        headerUnitView2.Rows[0].Cells["C33W_29L"].Value = s.C33W_29L;
                                        headerUnitView2.Rows[0].Cells["C33W_30L"].Value = s.C33W_30L;
                                        headerUnitView2.Rows[0].Cells["C33W_32L"].Value = s.C33W_32L;
                                        headerUnitView2.Rows[0].Cells["C33W_34L"].Value = s.C33W_34L;
                                        headerUnitView2.Rows[0].Cells["C34W_29L"].Value = s.C34W_29L;
                                        headerUnitView2.Rows[0].Cells["C34W_30L"].Value = s.C34W_30L;
                                        headerUnitView2.Rows[0].Cells["C34W_31L"].Value = s.C34W_31L;
                                        headerUnitView2.Rows[0].Cells["C34W_32L"].Value = s.C34W_32L;
                                        headerUnitView2.Rows[0].Cells["C34W_34L"].Value = s.C34W_34L;
                                        headerUnitView2.Rows[0].Cells["C36W_29L"].Value = s.C36W_29L;
                                        headerUnitView2.Rows[0].Cells["C36W_30L"].Value = s.C36W_30L;
                                        headerUnitView2.Rows[0].Cells["C36W_32L"].Value = s.C36W_32L;
                                        headerUnitView2.Rows[0].Cells["C36W_34L"].Value = s.C36W_34L;
                                        headerUnitView2.Rows[0].Cells["C38W_29L"].Value = s.C38W_29L;
                                        headerUnitView2.Rows[0].Cells["C38W_30L"].Value = s.C38W_30L;
                                        headerUnitView2.Rows[0].Cells["C38W_32L"].Value = s.C38W_32L;
                                        headerUnitView2.Rows[0].Cells["C38W_34L"].Value = s.C38W_34L;
                                        headerUnitView2.Rows[0].Cells["C40W_28L"].Value = s.C40W_28L;
                                        headerUnitView2.Rows[0].Cells["C40W_30L"].Value = s.C40W_30L;
                                        headerUnitView2.Rows[0].Cells["C40W_32L"].Value = s.C40W_32L;
                                        headerUnitView2.Rows[0].Cells["C40W_34L"].Value = s.C40W_34L;
                                        headerUnitView2.Rows[0].Cells["C42W_30L"].Value = s.C42W_30L;
                                        headerUnitView2.Rows[0].Cells["C42W_32L"].Value = s.C42W_32L;
                                        headerUnitView2.Rows[0].Cells["C42W_34L"].Value = s.C42W_34L;
                                        headerUnitView2.Rows[0].Cells["C44W_29L"].Value = s.C44W_29L;
                                        headerUnitView2.Rows[0].Cells["C44W_30L"].Value = s.C44W_30L;
                                        headerUnitView2.Rows[0].Cells["C44W_32L"].Value = s.C44W_32L;
                                        headerUnitView2.Rows[0].Cells["Sub_Total_"].Value = s.Sub_Total;

                                        //dt.Rows.Add(s.LOT_, s.STYLE_, s.ART, s.COLOR, s.COLORName, s.yaowei, s.C30W_29L, s.C30W_30L, s.C30W_32L, s.C31W_30L, s.C31W_32L, s.C32W_28L, s.C32W_30L, s.C32W_32L, s.C33W_29L, s.C33W_30L, s.C33W_32L, s.C33W_34L, s.C34W_29L, s.C34W_30L, s.C34W_31L, s.C34W_32L, s.C34W_34L, s.C36W_29L, s.C36W_30L, s.C36W_32L, s.C36W_34L, s.C38W_29L, s.C38W_30L, s.C38W_32L, s.C38W_34L, s.C40W_28L, s.C40W_30L, s.C40W_32L, s.C40W_34L, s.C42W_30L, s.C42W_32L, s.C42W_34L, s.C44W_29L, s.C44W_30L, s.C44W_32L, s.Sub_Total);
                                    }
                                    //headerUnitView2.DataSource = dt; 

                                #endregion
                                MessageBox.Show("读取成功！");
                            }
                            else 
                            {
                                MessageBox.Show("读取失败！原因:读取文件后缀非'xlsx");
                            }
                        }
                    }
                   
                }
            }
            catch (Exception ex) 
            {
                //throw ex;
                MessageBox.Show(ex.Message);
            }
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            for (int i = 0; i < 100; i++)
            {
                Thread.Sleep(10);
                worker.ReportProgress(i);
                if (worker.CancellationPending)  // 如果用户取消则跳出处理数据代码 
                {
                    e.Cancel = true;
                    break;
                }
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message);
            }
            else if (e.Cancelled)
            {
            }
            else
            {
            }
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void 删除ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult dr = MessageBox.Show("确定要删除选中信息？","信息",MessageBoxButtons.YesNo);
                if (dr == DialogResult.Yes)
                {
                    List<int> idtrr = new List<int>();
                    for (int i = this.dataGridView1.SelectedRows.Count; i > 0; i--)
                    {
                        if (dataGridView1.SelectedRows[i - 1].Cells[0].Value == null || dataGridView1.SelectedRows[i - 1].Cells[0].Value is DBNull)
                        {
                            DataRowView drv = dataGridView1.SelectedRows[i - 1].DataBoundItem as DataRowView;
                            if (drv != null)
                            {
                                drv.Delete();
                                i = i - 1;
                            }
                            i = i - 1;
                        }
                        else
                        {
                            idtrr.Add(Convert.ToInt32(dataGridView1.SelectedRows[i - 1].Cells[0].Value));

                        }
                    }
                    cal.deleteChiMa(idtrr);
                    this.backgroundWorker1.RunWorkerAsync();
                    JingDu form = new JingDu(this.backgroundWorker1, "删除中");// 显示进度条窗体
                    form.ShowDialog(this);
                    form.Close();
                    MessageBox.Show("删除成功！");
                    ChiMaDaPei_Load(sender, e);
                    //bindDataGirdview(comboBox1.Text);

                }
            }
            catch (Exception ex)
            {
                //throw ex;
                MessageBox.Show(ex.Message);
            }
        }
        public void bindDgvHeader() 
        {
            #region RGL1
                DataTable dt = new DataTable();
                dt.Columns.Add("面料", typeof(string));
                dt.Columns.Add("款式", typeof(string));
                dt.Columns.Add("货号", typeof(string));
                dt.Columns.Add("颜色", typeof(string));
                dt.Columns.Add("颜色编号", typeof(string));
                dt.Columns.Add("裤子", typeof(string));
                dt.Columns.Add("34R", typeof(string));
                dt.Columns.Add("36R", typeof(string));
                dt.Columns.Add("38R", typeof(string));
                dt.Columns.Add("40R", typeof(string));
                dt.Columns.Add("42R", typeof(string));
                dt.Columns.Add("44R", typeof(string));
                dt.Columns.Add("46R", typeof(string));
                dt.Columns.Add("48R", typeof(string));
                dt.Columns.Add("50R", typeof(string));
                dt.Columns.Add("52R", typeof(string));
                dt.Columns.Add("54R", typeof(string));
                dt.Columns.Add("56R", typeof(string));
                dt.Columns.Add("58R", typeof(string));
                dt.Columns.Add("60R", typeof(string));
                dt.Columns.Add("62R", typeof(string));
                dt.Columns.Add("36L", typeof(string));
                dt.Columns.Add("38L", typeof(string));
                dt.Columns.Add("40L", typeof(string));
                dt.Columns.Add("42L", typeof(string));
                dt.Columns.Add("44L", typeof(string));
                dt.Columns.Add("46L", typeof(string));
                dt.Columns.Add("48L", typeof(string));
                dt.Columns.Add("50L", typeof(string));
                dt.Columns.Add("52L", typeof(string));
                dt.Columns.Add("54L", typeof(string));
                dt.Columns.Add("56L", typeof(string));
                dt.Columns.Add("58L", typeof(string));
                dt.Columns.Add("60L", typeof(string));
                dt.Columns.Add("62L", typeof(string));
                dt.Columns.Add("34S", typeof(string));
                dt.Columns.Add("36S", typeof(string));
                dt.Columns.Add("38S", typeof(string));
                dt.Columns.Add("40S", typeof(string));
                dt.Columns.Add("42S", typeof(string));
                dt.Columns.Add("44S", typeof(string));
                dt.Columns.Add("46S", typeof(string));
                dt.Columns.Add("Sub Total: ", typeof(string));
                dataGridView1.DataSource = dt;
                DataGridViewHelper rowMergeView = new DataGridViewHelper(dataGridView1);
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(0, 1, "LOT#"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(1, 1, "STYLE"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(2, 1, "ART"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(3, 1, "COLOR"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(4, 1, "COLOR#"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(5, 1, "上衣"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(6, 1, "28"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(7, 1, "30"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(8, 1, "32"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(9, 1, "34"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(10, 1, "36"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(11, 1, "39"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(12, 1, "41"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(13, 1, "43"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(14, 1, "46"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(15, 1, "48"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(16, 1, "50"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(17, 1, "52"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(18, 1, "54"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(19, 1, "56"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(20, 1, "58"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(21, 1, "30"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(22, 1, "32"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(23, 1, "34"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(24, 1, "36"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(25, 1, "39"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(26, 1, "41"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(27, 1, "43"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(28, 1, "46"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(29, 1, "48"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(30, 1, "50"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(31, 1, "52"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(32, 1, "54"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(33, 1, "56"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(34, 1, "58"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(35, 1, "28"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(36, 1, "30"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(37, 1, "32"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(38, 1, "34"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(39, 1, "36"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(40, 1, "39"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(41, 1, "41"));
                rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(42, 1, "订单合计"));
            
            #endregion

            #region RGL2
            DataTable dt1= new DataTable();
            dt1.Columns.Add("面料", typeof(string));
            dt1.Columns.Add("款式", typeof(string));
            dt1.Columns.Add("货号", typeof(string));
            dt1.Columns.Add("颜色", typeof(string));
            dt1.Columns.Add("颜色编号", typeof(string));
            dt1.Columns.Add("裤子", typeof(string));
            dt1.Columns.Add("34R", typeof(string));
            dt1.Columns.Add("36R", typeof(string));
            dt1.Columns.Add("38R", typeof(string));
            dt1.Columns.Add("40R", typeof(string));
            dt1.Columns.Add("42R", typeof(string));
            dt1.Columns.Add("44R", typeof(string));
            dt1.Columns.Add("46R", typeof(string));
            dt1.Columns.Add("48R", typeof(string));
            dt1.Columns.Add("50R", typeof(string));
            dt1.Columns.Add("52R", typeof(string));
            dt1.Columns.Add("54R", typeof(string));
            dt1.Columns.Add("56R", typeof(string));
            dt1.Columns.Add("58R", typeof(string));
            dt1.Columns.Add("60R", typeof(string));
            dt1.Columns.Add("62R", typeof(string));
            dt1.Columns.Add("36L", typeof(string));
            dt1.Columns.Add("38L", typeof(string));
            dt1.Columns.Add("40L", typeof(string));
            dt1.Columns.Add("42L", typeof(string));
            dt1.Columns.Add("44L", typeof(string));
            dt1.Columns.Add("46L", typeof(string));
            dt1.Columns.Add("48L", typeof(string));
            dt1.Columns.Add("50L", typeof(string));
            dt1.Columns.Add("52L", typeof(string));
            dt1.Columns.Add("54L", typeof(string));
            dt1.Columns.Add("56L", typeof(string));
            dt1.Columns.Add("58L", typeof(string));
            dt1.Columns.Add("60L", typeof(string));
            dt1.Columns.Add("62L", typeof(string));
            dt1.Columns.Add("34S", typeof(string));
            dt1.Columns.Add("36S", typeof(string));
            dt1.Columns.Add("38S", typeof(string));
            dt1.Columns.Add("40S", typeof(string));
            dt1.Columns.Add("42S", typeof(string));
            dt1.Columns.Add("44S", typeof(string));
            dt1.Columns.Add("46S", typeof(string));
            dt1.Columns.Add("Sub Total: ", typeof(string));
            RGL2.DataSource = dt1;
            rowMergeView = new DataGridViewHelper(RGL2);

            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(0, 1, "LOT#"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(1, 1, "STYLE"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(2, 1, "ART"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(3, 1, "COLOR"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(4, 1, "COLOR#"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(5, 1, "上衣"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(6, 1, "28"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(7, 1, "30"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(8, 1, "32"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(9, 1, "34"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(10, 1, "36"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(11, 1, "38"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(12, 1, "40"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(13, 1, "42"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(14, 1, "44"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(15, 1, "46"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(16, 1, "48"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(17, 1, "50"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(18, 1, "52"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(19, 1, "54"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(20, 1, "56"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(21, 1, "30"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(22, 1, "32"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(23, 1, "34"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(24, 1, "36"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(25, 1, "38"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(26, 1, "40"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(27, 1, "42"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(28, 1, "44"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(29, 1, "46"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(30, 1, "48"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(31, 1, "50"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(32, 1, "52"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(33, 1, "54"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(34, 1, "56"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(35, 1, "28"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(36, 1, "30"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(37, 1, "32"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(38, 1, "34"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(39, 1, "36"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(40, 1, "38"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(41, 1, "40"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(42, 1, "订单合计"));

            #endregion

            #region SLIM
            DataTable dt3 = new DataTable();
            dt3.Columns.Add("面料", typeof(string));
            dt3.Columns.Add("款式", typeof(string));
            dt3.Columns.Add("货号", typeof(string));
            dt3.Columns.Add("颜色", typeof(string));
            dt3.Columns.Add("颜色编号", typeof(string));
            dt3.Columns.Add("裤子", typeof(string));
            dt3.Columns.Add("34R", typeof(string));
            dt3.Columns.Add("36R", typeof(string));
            dt3.Columns.Add("38R", typeof(string));
            dt3.Columns.Add("40R", typeof(string));
            dt3.Columns.Add("42R", typeof(string));
            dt3.Columns.Add("44R", typeof(string));
            dt3.Columns.Add("46R", typeof(string));
            dt3.Columns.Add("48R", typeof(string));
            dt3.Columns.Add("36L", typeof(string));
            dt3.Columns.Add("38L", typeof(string));
            dt3.Columns.Add("40L", typeof(string));
            dt3.Columns.Add("42L", typeof(string));
            dt3.Columns.Add("44L", typeof(string));
            dt3.Columns.Add("46L", typeof(string));
            dt3.Columns.Add("48L", typeof(string));
            dt3.Columns.Add("34S", typeof(string));
            dt3.Columns.Add("36S", typeof(string));
            dt3.Columns.Add("38S", typeof(string));
            dt3.Columns.Add("40S", typeof(string));
            dt3.Columns.Add("42S", typeof(string));
            dt3.Columns.Add("44S", typeof(string));
            dt3.Columns.Add("46S", typeof(string));
            dt3.Columns.Add("Sub Total: ", typeof(string));
            SLIM.DataSource = dt3;
            rowMergeView = new DataGridViewHelper(SLIM);

            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(0, 1, "LOT#"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(1, 1, "STYLE"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(2, 1, "ART"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(3, 1, "COLOR"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(4, 1, "COLOR#"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(5, 1, "上衣"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(6, 1, "28"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(7, 1, "30"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(8, 1, "32"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(9, 1, "34"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(10, 1, "36"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(11, 1, "38"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(12, 1, "40"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(13, 1, "42"));;
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(14, 1, "30"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(15, 1, "32"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(16, 1, "34"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(17, 1, "36"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(18, 1, "38"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(19, 1, "40"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(20, 1, "42"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(21, 1, "28"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(22, 1, "30"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(23, 1, "32"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(24, 1, "34"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(25, 1, "36"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(26, 1, "38"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(27, 1, "40"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(28, 1, "订单合计"));

            #endregion

            #region RGLJ
            DataTable dt4 = new DataTable();
            dt4.Columns.Add("面料", typeof(string));
            dt4.Columns.Add("款式", typeof(string));
            dt4.Columns.Add("货号", typeof(string));
            dt4.Columns.Add("颜色", typeof(string));
            dt4.Columns.Add("颜色编号", typeof(string));
            dt4.Columns.Add("上衣", typeof(string));
            dt4.Columns.Add("34R", typeof(string));
            dt4.Columns.Add("36R", typeof(string));
            dt4.Columns.Add("38R", typeof(string));
            dt4.Columns.Add("40R", typeof(string));
            dt4.Columns.Add("42R", typeof(string));
            dt4.Columns.Add("44R", typeof(string));
            dt4.Columns.Add("46R", typeof(string));
            dt4.Columns.Add("48R", typeof(string));
            dt4.Columns.Add("50R", typeof(string));
            dt4.Columns.Add("52R", typeof(string));
            dt4.Columns.Add("54R", typeof(string));
            dt4.Columns.Add("56R", typeof(string));
            dt4.Columns.Add("58R", typeof(string));
            dt4.Columns.Add("60R", typeof(string));
            dt4.Columns.Add("62R", typeof(string));
            dt4.Columns.Add("36L", typeof(string));
            dt4.Columns.Add("38L", typeof(string));
            dt4.Columns.Add("40L", typeof(string));
            dt4.Columns.Add("42L", typeof(string));
            dt4.Columns.Add("44L", typeof(string));
            dt4.Columns.Add("46L", typeof(string));
            dt4.Columns.Add("48L", typeof(string));
            dt4.Columns.Add("50L", typeof(string));
            dt4.Columns.Add("52L", typeof(string));
            dt4.Columns.Add("54L", typeof(string));
            dt4.Columns.Add("56L", typeof(string));
            dt4.Columns.Add("58L", typeof(string));
            dt4.Columns.Add("60L", typeof(string));
            dt4.Columns.Add("62L", typeof(string));
            dt4.Columns.Add("34S", typeof(string));
            dt4.Columns.Add("36S", typeof(string));
            dt4.Columns.Add("38S", typeof(string));
            dt4.Columns.Add("40S", typeof(string));
            dt4.Columns.Add("42S", typeof(string));
            dt4.Columns.Add("44S", typeof(string));
            dt4.Columns.Add("46S", typeof(string));
            dt4.Columns.Add("Sub Total: ", typeof(string));
            RGLJ.DataSource = dt4;
            //rowMergeView = new DataGridViewHelper(RGLJ);

            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(0, 1, "LOT#"));
            //rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(0, 1, "LOT1#"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(1, 1, "STYLE"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(2, 1, "ART"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(3, 1, "COLOR"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(4, 1, "COLOR#"));
            rowMergeView.Headers.Add(new DataGridViewHelper.TopHeader(43, 1, "订单合计"));

            #endregion


        }
        private void ChiMaDaPei_Load(object sender, EventArgs e)
        {
            try
            {
                //dataGridView1.TopRow.Cells[2].Text = "入库";
                //dataGridView1.TopRow.Cells[2].ColSpan = 2;


                //dataGridView1.TopRow.Cells[4].Text = "出库";
                //dataGridView1.TopRow.Cells[4].ColSpan = 2;
                bindDgvHeader();
                //comboBox1.SelectedIndexChanged -= comboBox1_SelectedIndexChanged;
                List<ChiMa_Dapeibiao> list = bindDataGirdview("").GroupBy(p => p.BiaoName).Select(pc => pc.First()).ToList<ChiMa_Dapeibiao>();
                //comboBox1.DisplayMember = "BiaoName";
                //comboBox1.ValueMember = "id";
                //comboBox1.DataSource = list;
                //comboBox1.SelectedIndexChanged += comboBox1_SelectedIndexChanged;
            }
            catch (Exception ex) 
            {
                //throw ex;
                MessageBox.Show(ex.Message);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //try
            //{

            //    this.backgroundWorker1.RunWorkerAsync(); // 运行 backgroundWorker 组件

            //    JingDu form = new JingDu(this.backgroundWorker1, "刷新中");// 显示进度条窗体
            //    form.ShowDialog(this);
            //    form.Close();
            //    //List<ChiMa_Dapeibiao> list = bindDataGirdview(comboBox1.Text);
            //    DataTable dt = new DataTable();
            //    dt.Columns.Add("id", typeof(int));
            //    for (int i = 0; i < dataGridView1.Columns.Count; i++)
            //    {
            //        if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("id"))
            //        {
            //            dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
            //        }
            //    }

            //    foreach (ChiMa_Dapeibiao s in list)
            //    {
            //        dt.Rows.Add(s.id, s.LOT__面料, s.STYLE_款式, s.ART_货号, s.COLOR_颜色, s.COLOR__颜色编号, s.JACKET_上衣_PANT_裤子, s.C34R_28, s.C36R_30, s.C38R_32, s.C40R___34, s.C42R_36, s.C44R_38, s.C46R_40, s.C48R_42, s.C50R_44, s.C52R_46, s.C54R_48, s.C56R_50, s.C58R_52, s.C60R_54, s.C62R_56, s.C36L_30, s.C38L_32, s.C40L_34, s.C42L_36, s.C44L_38, s.C46L_40, s.C48L_42, s.C50L_44, s.C52L_46, s.C54L_48, s.C56L_50, s.C58L_52, s.C60L_54, s.C62L_56, s.C34S_28, s.C36S_30, s.C38S_32, s.C40S_34, s.C42S_36, s.C44S_38, s.C46S_40, s.DingdanHeji);
            //    }
            //    dataGridView1.DataSource = dt;
            //}
            //catch (Exception ex)
            //{
            //    //throw ex;
            //    MessageBox.Show(ex.Message);
            //}
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    if (comboBox1.Text.Equals(string.Empty))
            //    {
            //        DataTable dt = new DataTable();
            //        dt.Columns.Add("id", typeof(int));
            //        for (int i = 0; i < dataGridView1.Columns.Count; i++)
            //        {
            //            if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("id"))
            //            {
            //                dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
            //            }
            //        }
            //        dataGridView1.DataSource = dt;
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);

            //}
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //try 
            //{
            //    DialogResult dr = MessageBox.Show("确定您要删除 名为'" + comboBox1.Text + "'的尺码搭配表吗？","信息",MessageBoxButtons.YesNo);
            //    if (dr == DialogResult.Yes)
            //    {
            //        this.backgroundWorker1.RunWorkerAsync();
            //        JingDu form = new JingDu(this.backgroundWorker1, "删除中");// 显示进度条窗体
            //        form.ShowDialog(this);
            //        form.Close();
            //        cal.deleteChiMaBiao(comboBox1.Text);
            //        MessageBox.Show("删除成功！");
            //    }
            //    ChiMaDaPei_Load(sender, e);
            //}
            //catch(Exception ex)
            //{
            //    //throw ex;
            //    MessageBox.Show(ex.Message);
            //}
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {

        }

        private void dataGridView2_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {

        }

        private void toolStripLabel1_Click_1(object sender, EventArgs e)
        {
            List<ChiMa_Dapeibiao> list = cal.SelectChiMaDapei("");
            List<RGL2> R = cal.SelectChiMaDapei2();
            List<SLIM> S = cal.SelectChiMaDapei3();
            List<RGLJ> RJ = cal.SelectChiMaDapei4();
            List<D_PANT> D = cal.SelectChiMaDapei5();
            List<C_PANT> C = cal.SelectChiMaDapei6();
            #region RGL1
            //list = cal.ReaderChiMaDapei(path);

            DataTable dt = new DataTable();
            //dt.Columns.Add("id", typeof(int));
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                if (!dataGridView1.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                {
                    dt.Columns.Add(dataGridView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                }
            }
            foreach (ChiMa_Dapeibiao s in list)
            {
                dt.Rows.Add( s.LOT__面料, s.STYLE_款式, s.ART_货号, s.COLOR_颜色, s.COLOR__颜色编号, s.JACKET_上衣_PANT_裤子, s.C34R_28, s.C36R_30, s.C38R_32, s.C40R___34, s.C42R_36, s.C44R_38, s.C46R_40, s.C48R_42, s.C50R_44, s.C52R_46, s.C54R_48, s.C56R_50, s.C58R_52, s.C60R_54, s.C62R_56, s.C36L_30, s.C38L_32, s.C40L_34, s.C42L_36, s.C44L_38, s.C46L_40, s.C48L_42, s.C50L_44, s.C52L_46, s.C54L_48, s.C56L_50, s.C58L_52, s.C60L_54, s.C62L_56, s.C34S_28, s.C36S_30, s.C38S_32, s.C40S_34, s.C42S_36, s.C44S_38, s.C46S_40, s.DingdanHeji);
            }
            dataGridView1.DataSource = dt;
            #endregion

            #region RGL2
            //R = cal.ReaderChiMaDapei2(path);
            dt = new DataTable();
            //dt.Columns.Add("id", typeof(int));
            for (int i = 0; i < RGL2.Columns.Count; i++)
            {
                if (!RGL2.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                {
                    dt.Columns.Add(RGL2.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                }
            }
            foreach (RGL2 s in R)
            {
                dt.Rows.Add(s.LOT_, s.STYLE_, s.ART, s.COLOR, s.COLORName, s.shangyi_kuzi, s.C34R, s.C36R, s.C38R, s.C40R, s.C42R, s.C44R, s.C46R, s.C48R, s.C50R, s.C52R, s.C54R, s.C56R, s.C58R, s.C60R, s.C62R, s.C36L, s.C38L, s.C40L, s.C42L, s.C44L, s.C46L, s.C48L, s.C50L, s.C52L, s.C54L, s.C56L, s.C58L, s.C60L, s.C62L, s.C34S, s.C36S, s.C38S, s.C40S, s.C42S, s.C44S, s.C46S, s.Sub_Total);
            }
            RGL2.DataSource = dt;

            #endregion

            #region SLIM
            //S = cal.ReaderChiMaDapei3(path);
            dt = new DataTable();
            //dt.Columns.Add("id", typeof(int));
            for (int i = 0; i < SLIM.Columns.Count; i++)
            {
                if (!SLIM.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                {
                    dt.Columns.Add(SLIM.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                }
            }
            foreach (SLIM s in S)
            {
                dt.Rows.Add(s.LOT_, s.STYLE_, s.ART, s.COLOR, s.COLOR_, s.shangyi_kuzi, s.C34R, s.C36R, s.C38R, s.C40R, s.C42R, s.C44R, s.C46R, s.C48R, s.C36L, s.C38L, s.C40L, s.C42L, s.C44L, s.C46L, s.C48L, s.C34S, s.C36S, s.C38S, s.C40S, s.C42S, s.C44S, s.C46S, s.Sub_Total);
            }
            SLIM.DataSource = dt;
            #endregion

            #region RGLJ
            //RJ = cal.ReaderChiMaDapei4(path);
            dt = new DataTable();
            //dt.Columns.Add("id", typeof(int));
            for (int i = 0; i < RGLJ.Columns.Count; i++)
            {
                if (!RGLJ.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                {
                    dt.Columns.Add(RGLJ.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                }
            }
            foreach (RGLJ s in RJ)
            {
                dt.Rows.Add(s.LOT_, s.STYLE_, s.ART, s.COLOR, s.COLOR_, s.shangyi, s.C34R, s.C36R, s.C38R, s.C40R, s.C42R, s.C44R, s.C46R, s.C48R, s.C50R, s.C52R, s.C54R, s.C56R, s.C58R, s.C60R, s.C62R, s.C36L, s.C38L, s.C40L, s.C42L, s.C44L, s.C46L, s.C48L, s.C50L, s.C52L, s.C54L, s.C56L, s.C58L, s.C60L, s.C62L, s.C34S, s.C36S, s.C38S, s.C40S, s.C42S, s.C44S, s.C46S, s.Sub_Total);
            }
            RGLJ.DataSource = dt;

            #endregion

            #region D.PANT
            //D = cal.ReaderChiMaDapei5(path);
            dt = new DataTable();
            //dt.Columns.Add("id", typeof(int));
            for (int i = 0; i < headerUnitView1.Columns.Count; i++)
            {
                if (!headerUnitView1.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                {
                    dt.Columns.Add(headerUnitView1.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                }
            }
            int RI = 0;
            foreach (D_PANT s in D)
            {
                headerUnitView1.Rows.Add();
                headerUnitView1.Rows[0].Cells["LOT"].Value = s.LOT_;
                headerUnitView1.Rows[0].Cells["STYLE"].Value = s.STYLE_;
                headerUnitView1.Rows[0].Cells["ART"].Value = s.ART;
                headerUnitView1.Rows[0].Cells["COLOR"].Value = s.COLOR;
                headerUnitView1.Rows[0].Cells["COLORName"].Value = s.COLORName;
                headerUnitView1.Rows[0].Cells["yaowei"].Value = s.yaowei;
                headerUnitView1.Rows[0].Cells["C30W_R_30L"].Value = s.C30W_R_30L;
                headerUnitView1.Rows[0].Cells["C30W_L_32L"].Value = s.C30W_L_32L;
                headerUnitView1.Rows[0].Cells["C32W_R_30L"].Value = s.C32W_R_30L;
                headerUnitView1.Rows[0].Cells["C32W_L_32L"].Value = s.C32W_L_32L;
                headerUnitView1.Rows[0].Cells["C34W_S_38L"].Value = s.C34W_S_38L;
                headerUnitView1.Rows[0].Cells["C34W_S_39L"].Value = s.C34W_S_39L;
                headerUnitView1.Rows[0].Cells["C34W_R_30L"].Value = s.C34W_R_30L;
                headerUnitView1.Rows[0].Cells["C34W_L_32L"].Value = s.C34W_L_32L;
                headerUnitView1.Rows[0].Cells["C34W_L_34L"].Value = s.C34W_L_34L;
                headerUnitView1.Rows[0].Cells["C36W_S_28L"].Value = s.C36W_S_28L;
                headerUnitView1.Rows[0].Cells["C36W_S_29L"].Value = s.C36W_S_29L;
                headerUnitView1.Rows[0].Cells["C36W_R_30L"].Value = s.C36W_R_30L;
                headerUnitView1.Rows[0].Cells["C36W_R_31L"].Value = s.C36W_R_31L;
                headerUnitView1.Rows[0].Cells["C38W_S_28L"].Value = s.C38W_S_28L;
                headerUnitView1.Rows[0].Cells["C38W_R_30L"].Value = s.C38W_R_30L;
                headerUnitView1.Rows[0].Cells["C38W_R_31L"].Value = s.C38W_R_31L;
                headerUnitView1.Rows[0].Cells["C38W_L_32L"].Value = s.C38W_L_32L;
                headerUnitView1.Rows[0].Cells["C38W_L_34L"].Value = s.C38W_L_34L;
                headerUnitView1.Rows[0].Cells["C40W_S_28L"].Value = s.C40W_S_28L;
                headerUnitView1.Rows[0].Cells["C40W_S_29L"].Value = s.C40W_S_29L;
                headerUnitView1.Rows[0].Cells["C40W_R_30L"].Value = s.C40W_R_30L;
                headerUnitView1.Rows[0].Cells["C40W_R_31L"].Value = s.C40W_R_31L;
                headerUnitView1.Rows[0].Cells["C40W_L_32L"].Value = s.C40W_L_32L;
                headerUnitView1.Rows[0].Cells["C40W_L_34L"].Value = s.C40W_L_34L;
                headerUnitView1.Rows[0].Cells["C42W_R_30L"].Value = s.C42W_R_30L;
                headerUnitView1.Rows[0].Cells["C42W_L_32L"].Value = s.C42W_L_32L;
                headerUnitView1.Rows[0].Cells["C42W_L_34L"].Value = s.C42W_L_34L;
                headerUnitView1.Rows[0].Cells["C44W_R_30L"].Value = s.C44W_R_30L;
                headerUnitView1.Rows[0].Cells["C44W_L_32L"].Value = s.C44W_L_32L;
                headerUnitView1.Rows[0].Cells["C44W_L_34L"].Value = s.C44W_L_34L;
                headerUnitView1.Rows[0].Cells["C46W_R_30L"].Value = s.C46W_R_30L;
                headerUnitView1.Rows[0].Cells["C46W_L_32L"].Value = s.C46W_L_32L;
                headerUnitView1.Rows[0].Cells["C48W_R_30L"].Value = s.C48W_R_30L;
                headerUnitView1.Rows[0].Cells["C48W_L_32L"].Value = s.C48W_L_32L;
                headerUnitView1.Rows[0].Cells["C50W_L_32L"].Value = s.C50W_L_32L;
                headerUnitView1.Rows[0].Cells["Sub_Total"].Value = s.Sub_Total;
                //headerUnitView1.Rows[0].Cells["ART"].Value = s.ART;
                //dt.Rows.Add(s.LOT_, s.STYLE_, s.ART, s.COLOR, s.COLORName, s.yaowei, s.C30W_R_30L, s.C30W_L_32L, s.C32W_R_30L, s.C32W_L_32L, s.C34W_S_38L, s.C34W_S_39L, s.C34W_R_30L, s.C34W_L_32L, s.C34W_L_34L, s.C36W_S_28L, s.C36W_S_29L, s.C36W_R_30L, s.C36W_R_31L, s.C38W_S_28L, s.C38W_R_30L, s.C38W_R_31L, s.C38W_L_32L, s.C38W_L_34L, s.C40W_S_28L, s.C40W_S_29L, s.C40W_R_30L, s.C40W_R_31L, s.C40W_L_32L, s.C40W_L_34L, s.C42W_R_30L, s.C42W_L_32L, s.C42W_L_34L, s.C44W_R_30L, s.C44W_L_32L, s.C44W_L_34L, s.C46W_R_30L, s.C46W_L_32L, s.C48W_R_30L, s.C48W_L_32L, s.C50W_L_32L, s.Sub_Total);
            }
            //headerUnitView1.DataSource = dt; 
            #endregion

            #region C.PANT
            //C = cal.ReaderChiMaDapei6(path);
            dt = new DataTable();
            //dt.Columns.Add("id", typeof(int));
            for (int i = 0; i < headerUnitView2.Columns.Count; i++)
            {
                if (!headerUnitView2.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                {
                    dt.Columns.Add(headerUnitView2.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                }
            }
            foreach (C_PANT s in C)
            {
                headerUnitView2.Rows.Add();
                headerUnitView2.Rows[0].Cells["LOT_"].Value = s.LOT_;
                headerUnitView2.Rows[0].Cells["STYLE_"].Value = s.STYLE_;
                headerUnitView2.Rows[0].Cells["ART_"].Value = s.ART;
                headerUnitView2.Rows[0].Cells["COLOR_"].Value = s.COLOR;
                headerUnitView2.Rows[0].Cells["COLORName_"].Value = s.COLORName;
                headerUnitView2.Rows[0].Cells["yaowei_"].Value = s.yaowei;
                headerUnitView2.Rows[0].Cells["C30W_29L"].Value = s.C30W_29L;
                headerUnitView2.Rows[0].Cells["C30W_30L"].Value = s.C30W_30L;
                headerUnitView2.Rows[0].Cells["C30W_32L"].Value = s.C30W_32L;
                headerUnitView2.Rows[0].Cells["C31W_30L"].Value = s.C31W_30L;
                headerUnitView2.Rows[0].Cells["C31W_32L"].Value = s.C31W_32L;
                headerUnitView2.Rows[0].Cells["C32W_28L"].Value = s.C32W_28L;
                headerUnitView2.Rows[0].Cells["C32W_30L"].Value = s.C32W_30L;
                headerUnitView2.Rows[0].Cells["C32W_32L"].Value = s.C32W_32L;
                headerUnitView2.Rows[0].Cells["C33W_29L"].Value = s.C33W_29L;
                headerUnitView2.Rows[0].Cells["C33W_30L"].Value = s.C33W_30L;
                headerUnitView2.Rows[0].Cells["C33W_32L"].Value = s.C33W_32L;
                headerUnitView2.Rows[0].Cells["C33W_34L"].Value = s.C33W_34L;
                headerUnitView2.Rows[0].Cells["C34W_29L"].Value = s.C34W_29L;
                headerUnitView2.Rows[0].Cells["C34W_30L"].Value = s.C34W_30L;
                headerUnitView2.Rows[0].Cells["C34W_31L"].Value = s.C34W_31L;
                headerUnitView2.Rows[0].Cells["C34W_32L"].Value = s.C34W_32L;
                headerUnitView2.Rows[0].Cells["C34W_34L"].Value = s.C34W_34L;
                headerUnitView2.Rows[0].Cells["C36W_29L"].Value = s.C36W_29L;
                headerUnitView2.Rows[0].Cells["C36W_30L"].Value = s.C36W_30L;
                headerUnitView2.Rows[0].Cells["C36W_32L"].Value = s.C36W_32L;
                headerUnitView2.Rows[0].Cells["C36W_34L"].Value = s.C36W_34L;
                headerUnitView2.Rows[0].Cells["C38W_29L"].Value = s.C38W_29L;
                headerUnitView2.Rows[0].Cells["C38W_30L"].Value = s.C38W_30L;
                headerUnitView2.Rows[0].Cells["C38W_32L"].Value = s.C38W_32L;
                headerUnitView2.Rows[0].Cells["C38W_34L"].Value = s.C38W_34L;
                headerUnitView2.Rows[0].Cells["C40W_28L"].Value = s.C40W_28L;
                headerUnitView2.Rows[0].Cells["C40W_30L"].Value = s.C40W_30L;
                headerUnitView2.Rows[0].Cells["C40W_32L"].Value = s.C40W_32L;
                headerUnitView2.Rows[0].Cells["C40W_34L"].Value = s.C40W_34L;
                headerUnitView2.Rows[0].Cells["C42W_30L"].Value = s.C42W_30L;
                headerUnitView2.Rows[0].Cells["C42W_32L"].Value = s.C42W_32L;
                headerUnitView2.Rows[0].Cells["C42W_34L"].Value = s.C42W_34L;
                headerUnitView2.Rows[0].Cells["C44W_29L"].Value = s.C44W_29L;
                headerUnitView2.Rows[0].Cells["C44W_30L"].Value = s.C44W_30L;
                headerUnitView2.Rows[0].Cells["C44W_32L"].Value = s.C44W_32L;
                headerUnitView2.Rows[0].Cells["Sub_Total_"].Value = s.Sub_Total;

                //dt.Rows.Add(s.LOT_, s.STYLE_, s.ART, s.COLOR, s.COLORName, s.yaowei, s.C30W_29L, s.C30W_30L, s.C30W_32L, s.C31W_30L, s.C31W_32L, s.C32W_28L, s.C32W_30L, s.C32W_32L, s.C33W_29L, s.C33W_30L, s.C33W_32L, s.C33W_34L, s.C34W_29L, s.C34W_30L, s.C34W_31L, s.C34W_32L, s.C34W_34L, s.C36W_29L, s.C36W_30L, s.C36W_32L, s.C36W_34L, s.C38W_29L, s.C38W_30L, s.C38W_32L, s.C38W_34L, s.C40W_28L, s.C40W_30L, s.C40W_32L, s.C40W_34L, s.C42W_30L, s.C42W_32L, s.C42W_34L, s.C44W_29L, s.C44W_30L, s.C44W_32L, s.Sub_Total);
            }
            //headerUnitView2.DataSource = dt; 

            #endregion

        }
    }
}
