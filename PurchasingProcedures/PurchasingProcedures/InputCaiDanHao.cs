using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using clsBuiness;
using logic;
namespace PurchasingProcedures
{
    public partial class InputCaiDanHao : Form
    {
        public clsAllnewLogic cal = new clsAllnewLogic();
        public string insertType;
        public string frmLabel;
        public Form f;
        public InputCaiDanHao(string type, string label, Form fm)
        {
            InitializeComponent();
            insertType = type;
            f = fm;
            frmLabel = label;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.backgroundWorker1.RunWorkerAsync();
            JingDu frm = new JingDu(this.backgroundWorker1, "生成裁单表中....");
            frm.ShowDialog();
            frm.Close();
            if (insertType.Equals("裁单"))
            {
                if (comboBox1.Text.Equals("RGL1"))
                {
                    CaiDan C = new CaiDan(textBox1.Text, comboBox1.Text, f);

                    if (!C.IsDisposed)
                    {
                        C.TopLevel = true;
                        C.MdiParent = f;
                        C.Show();
                    }
                }
                else if (comboBox1.Text.Equals("RGL2"))
                {
                    try
                    {
                        CaiDanRGL2 C = new CaiDanRGL2(textBox1.Text, comboBox1.Text, f);

                        if (!C.IsDisposed)
                        {
                            C.TopLevel = true;
                            C.MdiParent = f;
                            C.Show();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("异常：422363 " + ex);

                        return;
                        throw;
                    }

                }
                else if (comboBox1.Text.Equals("SLIM"))
                {
                    try
                    {
                        CaidanSLIM C = new CaidanSLIM(textBox1.Text, comboBox1.Text, f);

                        if (!C.IsDisposed)
                        {
                            C.TopLevel = true;
                            C.MdiParent = f;
                            C.Show();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("异常：20023 " + ex);

                        return;
                        throw;
                    }

                }
                else if (comboBox1.Text.Equals("RGLJ"))
                {
                    try
                    {
                        CaiDanRGLJ C = new CaiDanRGLJ(textBox1.Text, comboBox1.Text, f);

                        if (!C.IsDisposed)
                        {
                            C.TopLevel = true;
                            C.MdiParent = f;
                            C.Show();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("异常：457443 " + ex);

                        return;
                        throw;
                    }

                }
                else if (comboBox1.Text.Equals("D.PANT"))
                {
                    try
                    {
                        frmCaiDanD_PANT C = new frmCaiDanD_PANT(textBox1.Text, comboBox1.Text, f);

                        if (!C.IsDisposed)
                        {
                            C.TopLevel = true;
                            C.MdiParent = f;
                            C.Show();
                        }

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("异常：12314 " + ex);

                        return;
                        throw;
                    }
                }
                else if (comboBox1.Text.Equals("C.PANT"))
                {
                    try
                    {
                        frmCaiDanC C = new frmCaiDanC(textBox1.Text, comboBox1.Text, f);

                        if (!C.IsDisposed)
                        {
                            C.TopLevel = true;
                            C.MdiParent = f;
                            C.Show();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("异常：75564 " + ex);

                        return;

                        throw;
                    }

                }
                this.Close();
            }
        }

        private void InputCaiDanHao_Load(object sender, EventArgs e)
        {
            label1.Text = frmLabel;
            //List<ChiMa_Dapeibiao> list = cal.SelectChiMaDapei("").GroupBy(p => p.BiaoName).Select(pc => pc.First()).ToList<ChiMa_Dapeibiao>();
            comboBox1.DisplayMember = "BiaoName";
            comboBox1.ValueMember = "id";
            List<ChiMa_Dapeibiao> list = new List<ChiMa_Dapeibiao>();
            ChiMa_Dapeibiao cd = new ChiMa_Dapeibiao()
            {
                BiaoName = "RGL1",
                id = 1
            };
            ChiMa_Dapeibiao cd1 = new ChiMa_Dapeibiao()
            {
                BiaoName = "RGL2",
                id = 2
            };
            ChiMa_Dapeibiao cd2 = new ChiMa_Dapeibiao()
            {
                BiaoName = "SLIM",
                id = 3
            };
            ChiMa_Dapeibiao cd3 = new ChiMa_Dapeibiao()
            {
                BiaoName = "RGLJ",
                id = 4
            };
            ChiMa_Dapeibiao cd4 = new ChiMa_Dapeibiao()
            {
                BiaoName = "D.PANT",
                id = 5
            };
            ChiMa_Dapeibiao cd5 = new ChiMa_Dapeibiao()
            {
                BiaoName = "C.PANT",
                id = 6
            };
            list.Add(cd);
            list.Add(cd1);
            list.Add(cd2);
            list.Add(cd3);
            list.Add(cd4);
            list.Add(cd5);
            comboBox1.DisplayMember = "BiaoName";
            comboBox1.ValueMember = "id";
            comboBox1.DataSource = list;
            List<clsBuiness.KuanShiBiao> cdlist = cal.SelectKuanshi().GroupBy(g => g.STYLE).Select(pc => pc.First()).ToList<clsBuiness.KuanShiBiao>();
            textBox1.DataSource = cdlist;
            textBox1.DisplayMember = "STYLE";
            textBox1.ValueMember = "Id";

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
    }
}
