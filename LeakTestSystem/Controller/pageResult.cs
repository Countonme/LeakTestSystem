using Sunny.UI;
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

namespace Controller
{
    public partial class pageResult : Form
    {
        private readonly string _result;

        public pageResult(string result)
        {
            InitializeComponent();
            uiLabel1.ForeColor = Color.White;
            this.uiLabel1.Text = result;
            this.AcceptButton = null;
            this.TopMost = true;
            this.Load += PageResult_Load;
            _result = result;
        }

        private void PageResult_Load(object sender, EventArgs e)
        {
            SetColor(_result);
        }

        private void SetColor(string result)
        {
            switch (result)
            {
                case "PASS":
                    this.BackColor = Color.SpringGreen;
                    StartCloseTimer(1); // 1 秒后关闭
                    break;

                default:
                    this.BackColor = Color.Red;
                    StartCloseTimer(3); // 3 秒后关闭
                    break;
            }
        }

        private void StartCloseTimer(int seconds)
        {
            var timer = new System.Windows.Forms.Timer();
            timer.Interval = seconds * 1000;
            timer.Tick += (s, e) =>
            {
                timer.Stop();
                this.Close();
            };
            timer.Start();
        }
    }
}