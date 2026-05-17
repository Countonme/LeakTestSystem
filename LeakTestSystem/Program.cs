using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Controller;
using LeakTestSystem.Controller;

namespace LeakTestSystem
{
    internal static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        private static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new pageResult("PASS"));
            //Application.Run(new FrmSN(6,3,false));
            Application.Run(new FrmMaster());
        }
    }
}