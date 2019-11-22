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
using System.IO;
using Spire.Xls;
namespace PurchasingProcedures
{
   
    public partial class shengchengBiaoge : Form
    {
        private string cdNo;

        protected GongNeng2 gn;
        protected clsAllnewLogic cal;
        private string foldPath;
        private int lie;
        private string STYLE;
        private List<string> color;
        private string insertStr;
        public shengchengBiaoge(string caidanNo)
        {
            InitializeComponent();
            cdNo = caidanNo;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            gn = new GongNeng2();
            cal = new clsAllnewLogic();
            List<clsBuiness.CaiDan> cdlist = gn.selectCaiDan("").GroupBy(c => c.LOT).Select(s => s.First()).ToList<clsBuiness.CaiDan>();
            STYLE = cdlist[0].STYLE;

        }

        private void shengchengBiaoge_Load(object sender, EventArgs e)
        {
            //配色
             CreatePeiSe();
            //单耗
             CreateDanHao();
            
        }

        private  void CreateDanHao()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("面料", typeof(string));
            dt.Columns.Add("货号", typeof(string));
            dt.Columns.Add("幅宽", typeof(string));
            dt.Columns.Add("色号&颜色", typeof(string));
            dt.Columns.Add("单耗", typeof(string));
            dt.Columns.Add("单价", typeof(string));
            dt.Columns.Add("金额", typeof(string));
            List<DanHao> dhlist = cal.SelectDanHao("").Where(c => c.CaiDanNo.Trim().ToUpper().Equals(cdNo.Trim().ToUpper())).ToList<DanHao>() ;
            foreach(DanHao dh in dhlist)
            {
                dt.Rows.Add(dh.Name,dh.HuoHao,dh.GuiGe,dh.Yanse,dh.DanHao1,dh.Danjia,dh.Jine);
            }
            dgv_dh.DataSource = dt;
        }
        private void CreatePeiSe()
        {
            color = new List<string>();
            lie = 0;
            List<clsBuiness.CaiDan> cdlist = gn.selectCaiDan("").GroupBy(c => c.LOT).Select(s => s.First()).ToList<clsBuiness.CaiDan>();
            DataTable dt = new DataTable();
            dt.Columns.Add("品名", typeof(string));
            dt.Columns.Add("货号", typeof(string));
            dt.Columns.Add("规格/幅宽", typeof(string));
            string dgvSTR = "";
            foreach (clsBuiness.CaiDan cd in cdlist)
            {
                dt.Columns.Add(cd.LOT, typeof(string));
                dgvSTR =dgvSTR+"="+ cd.LOT;
            }
            List<clsBuiness.CaiDan> listcd = cdlist.FindAll(c => c.CaiDanHao.Equals(cdNo)) ;
            string mlNo ="";
            if (listcd != null)
            {
                mlNo = listcd[0].MianLiao;
            }
            List<clsBuiness.PeiSe> ps = cal.selectPeise("").FindAll(p => p.Fabrics.Trim().ToUpper().Equals(mlNo));
            insertStr = "";
            foreach (clsBuiness.PeiSe p in ps) 
            {
                insertStr = p.PingMing+"="+ p.HuoHao+"="+ p.GuiGe;
                lie = 0;
                if (dgvSTR.Contains("61601C1")) 
                {
                    insertStr = insertStr + "=" + p.C61601C1;
                    color.Add("黑 BLACK");
                    lie++;
                }
                if (dgvSTR.Contains("61602C1"))
                {
                    insertStr = insertStr + "=" + p.C61602C1;
                    color.Add("碳灰CHARCOAL");
                    lie++;
                }
                if (dgvSTR.Contains("61603C1"))
                {
                    insertStr = insertStr + "=" + p.C61603C1;
                    color.Add("海军蓝 NAVY");
                    lie++;
                }
                if (dgvSTR.Contains("61605C1"))
                {
                    insertStr = insertStr + "=" + p.C61605C1;
                    color.Add("米色 TAN");
                    lie++;
                }
                if (dgvSTR.Contains("61606C1"))
                {
                    insertStr = insertStr + "=" + p.C61606C1;
                    color.Add("灰色 GREY");
                    lie++;
                }
                if (dgvSTR.Contains("61607C1"))
                {
                    insertStr = insertStr + "=" + p.C61607C1;
                    color.Add("银灰色 SILVER GREY");
                    lie++;
                }
                if (dgvSTR.Contains("61609C1"))
                {
                    insertStr = insertStr + "=" + p.C61609C1;
                    color.Add("棕色 BROWN");
                    lie++;
                }
                if (dgvSTR.Contains("61611C1"))
                {
                    insertStr = insertStr + "=" + p.C61601C1;
                    color.Add("枣红色 BURGUNDY");
                    lie++;
                }
                if (dgvSTR.Contains("61618C1"))
                {
                    insertStr = insertStr + "=" + p.C61618C1;
                    color.Add("蓝色 BLUE");
                    lie++;
                }
                if (dgvSTR.Contains("61624C1"))
                {
                    insertStr = insertStr + "=" + p.C61624C1;
                    color.Add("橄榄绿 OLIVE");
                    lie++;
                }
                if (dgvSTR.Contains("61627C1"))
                {
                    insertStr = insertStr + "=" + p.C61627C1;
                    color.Add("钴蓝色 COBALT BLUE");
                    lie++;
                }
                if (dgvSTR.Contains("61631C1"))
                {
                    insertStr = insertStr + "=" + p.C61631C1;
                    color.Add("紫罗兰 IMPERIAL PURPLE");
                    lie++;
                }
                if (dgvSTR.Contains("61632C1"))
                {
                    insertStr = insertStr + "=" + p.C61632C1;
                    color.Add("宝石蓝 ROYAL BLUE");
                    lie++;
                }
                if (dgvSTR.Contains("61633C1"))
                {
                    insertStr = insertStr + "=" + p.C61633C1;
                    color.Add("苋红 RUMBA RED");
                    lie++;
                }
                if (dgvSTR.Contains("61634C1"))
                {
                    insertStr = insertStr + "=" + p.C61634C1;
                    color.Add("鹿褐色 GOLDEN BROWN");
                    lie++;
                }

                dt.Rows.Add(insertStr.Split('='));
            }
            
            dgv_ps.DataSource = dt;

        }
      
        private void button3_Click(object sender, EventArgs e)
        {
           //try 
           // {
            ShengChengBiaoGeXuanZe scbgz = new ShengChengBiaoGeXuanZe("打印", dgv_ps,dgv_dh,color,lie,STYLE,cdNo);
                scbgz.ShowDialog();

                 
                    
                    

                
           // }
           //catch (Exception ex) { throw ex; }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("确认要保存‘单耗’‘配色’两份表格吗？", "系统提示", MessageBoxButtons.YesNo);
            if (dr == DialogResult.Yes)
            {
                try
                {
                    FolderBrowserDialog dialog = new FolderBrowserDialog();
                    dialog.Description = "请选择文件路径";
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        foldPath = dialog.SelectedPath;
                        CreateExcel(foldPath);
                        MessageBox.Show("保存成功！");
                    }
                }
                catch (Exception ex) { throw ex; }
            }
            else
            {
                ShengChengBiaoGeXuanZe scbgz = new ShengChengBiaoGeXuanZe("保存",dgv_ps,dgv_dh,color,lie,STYLE,cdNo);
                scbgz.ShowDialog();

            }
        }



        private void CreateExcel(string path)
        {
            DataTable dt = new DataTable();
           
            for (int i = 0; i < dgv_ps.Columns.Count; i++)
            {
                if (!dgv_ps.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                {
                    dt.Columns.Add(dgv_ps.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                }
            }




            string C1str = "品名=货号=规格/幅宽";
            for (int i = 3; i < dgv_ps.ColumnCount; i++)
            {

                C1str = C1str + "=" + dgv_ps.Columns[i].HeaderCell.Value.ToString();
            }
            dt.Rows.Add(C1str.Split('='));
            string C2str  = "面料颜色= = ";
            for (int i = 0; i < color.Count; i++)
            {
                if (!C2str.Contains(color[i]))
                {
                    C2str = C2str + "=" + color[i];
                }
            }
            dt.Rows.Add(C2str.Split('='));
           
               
                for (int i = 0; i < dgv_ps.Rows.Count; i++)
                {
                    string str = "";
                //if (dgv_ps.Rows[i].Cells[6].Value != null)
                //{
                    str = dgv_ps.Rows[i].Cells[0].Value + "=" + dgv_ps.Rows[i].Cells[1].Value + "=" + dgv_ps.Rows[i].Cells[2].Value;

                    for (int j = 3; j < lie+3; j++)
                    {
                    
                        if (dgv_ps.Rows[i].Cells[j].Value!=null )
                        {
                            str = str + "=" + dgv_ps.Rows[i].Cells[j].Value;
                        }
                        
                    }
                    dt.Rows.Add(str.Split('='));
                //}
            }



            DataTable dt2 = new DataTable();

            for (int i = 0; i < dgv_dh.Columns.Count; i++)
            {
                if (!dgv_dh.Columns[i].HeaderCell.Value.ToString().Equals("id"))
                {
                    dt2.Columns.Add(dgv_dh.Columns[i].HeaderCell.Value.ToString(), typeof(String));
                }
            }
            for (int i = 0; i < dgv_dh.Rows.Count; i++)
            {
                if (dgv_dh.Rows[i].Cells[6].Value != null)
                {
                    string str = "";
                    str = dgv_dh.Rows[i].Cells[0].Value + "=" + dgv_dh.Rows[i].Cells[1].Value + "=" + dgv_dh.Rows[i].Cells[2].Value + "=" + dgv_dh.Rows[i].Cells[3].Value + "=" + dgv_dh.Rows[i].Cells[4].Value + "=" + dgv_dh.Rows[i].Cells[5].Value + "=" + dgv_dh.Rows[i].Cells[6].Value;
                    dt2.Rows.Add(str.Split('='));
                }
            }
            gn.SavePeiSeToExcel(dt,dt2, path,STYLE ,cdNo);
            foldPath = path + "\\配色表-" + STYLE + "-" + cdNo + ".xls";
            //foldPath2 = path + "\\单耗-" + STYLE + "-" + cdNo + ".xls";
        }




    }
}
