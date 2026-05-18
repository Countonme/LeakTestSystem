using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Controller;
using LeakTestSystem.Controller;
using OfficeOpenXml;

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
            ExcelPackage.License.SetNonCommercialPersonal("LeakTestSystem");
            //Application.Run(new pageResult("PASS"));
            //Application.Run(new FrmSN(6, 3, false));
            Application.Run(new FrmMaster());
        }
    }
}