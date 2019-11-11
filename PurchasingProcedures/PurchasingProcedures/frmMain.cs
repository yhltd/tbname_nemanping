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
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();

        }

        private void frmMain_Load(object sender, EventArgs e)
        {

        }

        private void toolStripLabel1_Click(object sender, EventArgs e)
        {

        }

        private void 手工ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void 色号表录入ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SeHaoBiaoLuru shlr = new SeHaoBiaoLuru();
            shlr.MdiParent = this;
            shlr.Show();
        }

        private void 尺码搭配ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ChiMaDaPei cmi = new ChiMaDaPei();
            cmi.MdiParent = this;
            cmi.Show();
        }

        private void 色号录入ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            SeHaoBiaoLuru shlr = new SeHaoBiaoLuru();
            if (!HaveOpened(this, shlr.Name))
            {
                shlr.MdiParent = this;
                shlr.Show();
            }
            else
            {
                shlr.TopMost = true;
            }
            
        }
        private bool HaveOpened(Form _monthForm, string _childrenFormName)
        {
            //查看窗口是否已经被打开
            bool bReturn = false;
            for (int i = 0; i < _monthForm.MdiChildren.Length; i++)
            {
                if (_monthForm.MdiChildren[i].Name == _childrenFormName)
                {
                    _monthForm.MdiChildren[i].BringToFront();//将控件带到 Z 顺序的前面。
                    bReturn = true;
                    break;
                }
            }
            return bReturn;
        }
        private void 尺码搭配表录入ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ChiMaDaPei cmi = new ChiMaDaPei();
            if (!HaveOpened(this, cmi.Name))
            {
                cmi.MdiParent = this;
                cmi.Show();
            }
            else
            {
                cmi.TopMost = true;
            }
            
        }

        private void 款式表数据录入ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            KuanShibiao ks = new KuanShibiao();
            if (!HaveOpened(this, ks.Name))
            {
                ks.MdiParent = this;
                ks.Show();
            }
            else
            {
                ks.TopMost = true;
            }
            
            
        }

        private void 单号表录入ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DanHaoFrm dh = new DanHaoFrm();
            if (!HaveOpened(this, dh.Name))
            {
                dh.MdiParent = this;
                dh.Show();
            }
            else
            {
                dh.TopMost = true;
            }
            
        }

        private void 面料表录入ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PeiSeBiaoLuru psbl = new PeiSeBiaoLuru();
            if (!HaveOpened(this, psbl.Name))
            {
                psbl.MdiParent = this;
                psbl.Show();
            }
            else
            {
                psbl.TopMost = true;
            }
            
        }

        private void 加工厂表录入ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Factoryinput fci = new Factoryinput();
            if (!HaveOpened(this, fci.Name))
            {
                fci.MdiParent = this;
                fci.Show();
            }
            else
            {
                fci.TopMost = true;
            }
        }

        private void 库存表录入ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Kucun kc = new Kucun();
            if (!HaveOpened(this, kc.Name))
            {
                kc.MdiParent = this;
                kc.Show();
            }
            else
            {
                kc.TopMost = true;
            }
        }
    }
}
