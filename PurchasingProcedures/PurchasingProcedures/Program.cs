﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PurchasingProcedures
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //var hu = new HeaderUnitView();
            var form = new Login();
            if (form.ShowDialog() != DialogResult.OK)
            {
                Application.Exit();

            }
            else
            {
                Application.Run(new frmMain());
                //Application.Run(new Form1());
            }
            
        }
    }
}
