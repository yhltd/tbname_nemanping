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
    public partial class ETD_ETA_LDP : Form
    {
        private string caidanhao;
        private string jiagongchang;
        private string style;
        protected clsAllnewLogic cal;
        protected GongNeng2 gn ;
        private List<HeSuan> hs;
        public ETD_ETA_LDP(string cdNo ,string jgc,string kuanshi,List<HeSuan> hesuan )
        {
            caidanhao = cdNo;
            jiagongchang = jgc;
            style = kuanshi;
            hs = hesuan;
            cal = new clsAllnewLogic();
            gn = new GongNeng2();
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (!textBox6.Text.Equals(string.Empty) && textBox1.Text != "")
            {
                textBox26.Text = (Convert.ToDouble(textBox6.Text) / 1.13 / Convert.ToDouble(textBox1.Text) * 0.01).ToString();
            }
            if (!textBox7.Text.Equals(string.Empty) && textBox1.Text != "")
            {
                textBox28.Text = (Convert.ToDouble(textBox7.Text) / 1.13 / Convert.ToDouble(textBox1.Text)).ToString();
            }
            if (!textBox8.Text.Equals(string.Empty) && textBox1.Text != "")
            {
                textBox30.Text = (Convert.ToDouble(textBox8.Text) / 1.13 / Convert.ToDouble(textBox1.Text)).ToString();
            }
        }
        private bool IsNumberic(string oText)
        {
            try
            {
                int var1 = Convert.ToInt32(oText);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void ETD_ETA_LDP_Load(object sender, EventArgs e)
        {
            List<clsBuiness.DanHao> danhao = cal.SelectDanHao("").FindAll(f => f.Style.Trim().Contains(style.Trim()) && f.JiaGongChang.Trim().Equals(jiagongchang.Trim()) && f.CaiDanNo.Trim().Equals(caidanhao.Trim()));
            if (danhao.Count > 0)
            {
                textBox6.Text = danhao.Average(d => Convert.ToDouble(d.Jine)).ToString();
                textBox7.Text = danhao.Sum(sc => Convert.ToDouble(sc.Jine)).ToString();

                textBox16.Text = hs.Sum(s => Convert.ToInt32(s.实际出口数量)).ToString();
            }
            else 
            {
                MessageBox.Show("生成失败！原因：找不到预计成本/实际成本单！");
                this.Close();
                return;
            }
        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            if (!textBox7.Text.Equals(string.Empty) && textBox1.Text != "")
            {
                textBox28.Text = (Convert.ToDouble(textBox7.Text) / 1.13 / Convert.ToDouble(textBox1.Text)).ToString();
            }
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            if (!textBox8.Text.Equals(string.Empty) && textBox1.Text != "")
            {
                textBox30.Text = (Convert.ToDouble(textBox8.Text) / 1.13 / Convert.ToDouble(textBox1.Text)).ToString();
            }
        }

        private void textBox16_TextChanged(object sender, EventArgs e)
        {
            if (!textBox26.Text.Equals(string.Empty) && textBox16.Text != "")
            {
                textBox36.Text = (Convert.ToDouble(textBox26.Text) * Convert.ToDouble(textBox16.Text)).ToString();
            }
            if (!textBox28.Text.Equals(string.Empty) && textBox16.Text != "")
            {
                textBox38.Text = (Convert.ToDouble(textBox28.Text) * Convert.ToDouble(textBox16.Text)).ToString();
            }
            if (!textBox30.Text.Equals(string.Empty) && textBox16.Text != "")
            {
                textBox40.Text = (Convert.ToDouble(textBox30.Text) * Convert.ToDouble(textBox16.Text)).ToString();
            }
        }

        private void textBox26_TextChanged(object sender, EventArgs e)
        {
            if (!textBox26.Text.Equals(string.Empty) && textBox16.Text != "")
            {
                textBox36.Text = (Convert.ToDouble(textBox26.Text) * Convert.ToDouble(textBox16.Text)).ToString();
            }
            if (!textBox30.Text.Equals(string.Empty) && !textBox28.Text.Equals(string.Empty) && !textBox26.Text.Equals(string.Empty) && !textBox23.Text.Equals(string.Empty))
            {
                textBox34.Text = ((Convert.ToDouble(textBox30.Text) + Convert.ToDouble(textBox28.Text) + Convert.ToDouble(textBox26.Text)) * (Convert.ToDouble(textBox23.Text) + 1)).ToString();
            }
        }

        private void textBox28_TextChanged(object sender, EventArgs e)
        {
            if (!textBox28.Text.Equals(string.Empty) && textBox16.Text != "")
            {
                textBox38.Text = (Convert.ToDouble(textBox28.Text) * Convert.ToDouble(textBox16.Text)).ToString();
            }
            if (!textBox30.Text.Equals(string.Empty) && !textBox28.Text.Equals(string.Empty) && !textBox26.Text.Equals(string.Empty) && !textBox23.Text.Equals(string.Empty))
            {
                textBox34.Text = ((Convert.ToDouble(textBox30.Text) + Convert.ToDouble(textBox28.Text) + Convert.ToDouble(textBox26.Text)) * (Convert.ToDouble(textBox23.Text) + 1)).ToString();
            }
        }

        private void textBox30_TextChanged(object sender, EventArgs e)
        {
            if (!textBox30.Text.Equals(string.Empty) && textBox16.Text != "")
            {
                textBox40.Text = (Convert.ToDouble(textBox30.Text) * Convert.ToDouble(textBox16.Text)).ToString();
            }
            if (!textBox30.Text.Equals(string.Empty) && !textBox28.Text.Equals(string.Empty) && !textBox26.Text.Equals(string.Empty) && !textBox23.Text.Equals(string.Empty))
            {
                textBox34.Text = ((Convert.ToDouble(textBox30.Text) + Convert.ToDouble(textBox28.Text) + Convert.ToDouble(textBox26.Text)) * (Convert.ToDouble(textBox23.Text)+1)).ToString();
            }
        }

        private void textBox36_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void textBox34_TextChanged(object sender, EventArgs e)
        {
            if (!textBox34.Text.Equals(string.Empty) && !textBox23.Text.Equals(string.Empty))
            {
                textBox44.Text = (Convert.ToDouble(textBox34.Text) * Convert.ToDouble(textBox23.Text )).ToString();
            }
            if (!textBox34.Text.Equals(string.Empty) && !textBox18.Text.Equals(string.Empty))
            {
                textBox29.Text = (Convert.ToDouble(textBox34.Text) + Convert.ToDouble(textBox18.Text)).ToString();
            }
        }

        private void textBox19_TextChanged(object sender, EventArgs e)
        {
            if (!textBox19.Text.Equals(string.Empty) && !textBox16.Text.Equals(string.Empty))
            {
                textBox18.Text = (Convert.ToDouble(textBox19.Text) + Convert.ToDouble(textBox16.Text)).ToString();
            }
            if (!textBox44.Text.Equals(string.Empty) && !textBox19.Text.Equals(string.Empty))
            {
                textBox39.Text = (Convert.ToDouble(textBox44.Text) + Convert.ToDouble(textBox19.Text)).ToString();
            }
            if (!textBox39.Text.Equals(string.Empty) && !textBox19.Text.Equals(string.Empty))
            {
                textBox24.Text = (Convert.ToDouble(textBox39.Text) + Convert.ToDouble(textBox19.Text) ).ToString();
            }
        }

        private void textBox18_TextChanged(object sender, EventArgs e)
        {
            if (!textBox34.Text.Equals(string.Empty) && !textBox18.Text.Equals(string.Empty))
            {
                textBox29.Text = (Convert.ToDouble(textBox34.Text) + Convert.ToDouble(textBox18.Text)).ToString();
            }
        }

        private void textBox44_TextChanged(object sender, EventArgs e)
        {
            if (!textBox44.Text.Equals(string.Empty) && !textBox19.Text.Equals(string.Empty))
            {
                textBox39.Text = (Convert.ToDouble(textBox44.Text) + Convert.ToDouble(textBox19.Text)).ToString();
            }
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            if (!textBox29.Text.Equals(string.Empty) && !textBox22.Text.Equals(string.Empty) && !textBox5.Text.Equals(string.Empty) && !textBox16.Text.Equals(string.Empty))
            {
                textBox35.Text = (Convert.ToDouble(textBox29.Text) + Convert.ToDouble(textBox22.Text) + (Convert.ToDouble(textBox5.Text) / Convert.ToDouble(textBox16.Text))).ToString();
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox21_TextChanged(object sender, EventArgs e)
        {
            if (!textBox21.Text.Equals(string.Empty) && !textBox25.Text.Equals(string.Empty))
            {
                textBox22.Text = (Convert.ToDouble(textBox21.Text) *(Convert.ToDouble(textBox25.Text)*0.01)).ToString();
            }
        }

        private void textBox25_TextChanged(object sender, EventArgs e)
        {
            if (!textBox21.Text.Equals(string.Empty) && !textBox25.Text.Equals(string.Empty))
            {
                textBox22.Text = (Convert.ToDouble(textBox21.Text) * (Convert.ToDouble(textBox25.Text) * 0.01)).ToString();
            }
        }

        private void textBox29_TextChanged(object sender, EventArgs e)
        {
            if (!textBox29.Text.Equals(string.Empty) && !textBox22.Text.Equals(string.Empty) && !textBox5.Text.Equals(string.Empty) && !textBox16.Text.Equals(string.Empty))
            {
                textBox35.Text = (Convert.ToDouble(textBox29.Text) + Convert.ToDouble(textBox22.Text) + (Convert.ToDouble(textBox5.Text) / Convert.ToDouble(textBox16.Text))).ToString();
            }
        }

        private void textBox22_TextChanged(object sender, EventArgs e)
        {
            if (!textBox29.Text.Equals(string.Empty) && !textBox22.Text.Equals(string.Empty) && !textBox5.Text.Equals(string.Empty) && !textBox16.Text.Equals(string.Empty))
            {
                textBox35.Text = (Convert.ToDouble(textBox29.Text) + Convert.ToDouble(textBox22.Text) + (Convert.ToDouble(textBox5.Text) / Convert.ToDouble(textBox16.Text))).ToString();
            }
        }

        private void textBox39_TextChanged(object sender, EventArgs e)
        {
            if (!textBox39.Text.Equals(string.Empty) && !textBox24.Text.Equals(string.Empty) && !textBox5.Text.Equals(string.Empty) )
            {
                textBox45.Text = (Convert.ToDouble(textBox39.Text) + Convert.ToDouble(textBox24.Text) + Convert.ToDouble(textBox5.Text)).ToString();
            }
        }

        private void textBox23_TextChanged(object sender, EventArgs e)
        {
            if (!textBox30.Text.Equals(string.Empty) && !textBox28.Text.Equals(string.Empty) && !textBox26.Text.Equals(string.Empty) && !textBox23.Text.Equals(string.Empty))
            {
                textBox34.Text = ((Convert.ToDouble(textBox30.Text) + Convert.ToDouble(textBox28.Text) + Convert.ToDouble(textBox26.Text)) * (Convert.ToDouble(textBox23.Text) + 1)).ToString();
            }
        }
    }
}
