using LeakTestSystem.Model;
using LeakTestSystem.Services;
using LeakTestSystem.Services.MES;
using Sunny.UI;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace LeakTestSystem
{
    public partial class FrmMaster : UIForm
    {
        public ModbusIoController modbusIo;
        private SerialPort serialPort1, serialPort2, serialPort3, serialPort4, serialPort5, serialPort6, serialPort7;
        private bool flagProductionModel;
        private List<ScanModel> snList = new List<ScanModel>();
        private int maxCount = 6;

        public FrmMaster()
        {
            InitializeComponent();
            this.Load += Form1_Load;
            this.switchCom7.Click += SwitchCom7_Click;
            this.IntegerUpDownChannels.TextChanged += IntegerUpDownChannels_TextChanged;
            //this.FrmMaster_FormClosing += FrmMaster_FormClosing;
        }

        private void FrmMaster_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.ShowAskDialog("您确定退出窗体吗", true);
            {
            }
        }

        private void IntegerUpDownChannels_TextChanged(object sender, EventArgs e)
        {
            if (!switchCom7.Active)
            {
                this.titlePanel8.Text = $"Master Channels ({IntegerUpDownChannels.Value})";
                this.ShowSuccessTip(IntegerUpDownChannels.Value.ToString());
                initTitlePanelColor();
                maxCount = IntegerUpDownChannels.Value;
            }
            else
            {
                this.ShowErrorNotifier("请先关闭通道在修改通道数量");
            }
        }

        private void SwitchCom7_Click(object sender, EventArgs e)
        {
            if (switchCom7.Active)
            {
                // 1. 判断是否存在 COM1~COM7
                if (!CheckCom1To7Exists())
                {
                    MessageBox.Show("未检测到完整 COM1~COM7");
                    switchCom7.Active = false;
                    return;
                }

                // 2. 初始化
                InitAllPorts();

                // 3. 打开所有串口
                OpenAllPorts();

                // 4. 重点：COM7 初始化 Modbus
                modbusIo = new ModbusIoController(serialPort7);

                MessageBox.Show("所有串口已打开，Modbus已就绪");
            }
            else
            {
                CloseAllPorts();
                MessageBox.Show("所有串口已关闭");
            }
        }

        private bool CheckCom1To7Exists()
        {
            var comList = SerialPort.GetPortNames();

            for (int i = 1; i <= 7; i++)
            {
                if (!comList.Contains($"COM{i}"))
                {
                    return false;
                }
            }

            return true;
        }

        private void InitAllPorts()
        {
            serialPort1 = new SerialPort("COM1", 9600);
            serialPort2 = new SerialPort("COM2", 9600);
            serialPort3 = new SerialPort("COM3", 9600);
            serialPort4 = new SerialPort("COM4", 9600);
            serialPort5 = new SerialPort("COM5", 9600);
            serialPort6 = new SerialPort("COM6", 9600);

            // COM7 = Modbus RTU
            serialPort7 = new SerialPort("COM7", 9600);
            // 统一绑定事件（关键）
            serialPort1.DataReceived += Serial_DataReceived;
            serialPort2.DataReceived += Serial_DataReceived;
            serialPort3.DataReceived += Serial_DataReceived;
            serialPort4.DataReceived += Serial_DataReceived;
            serialPort5.DataReceived += Serial_DataReceived;
            serialPort6.DataReceived += Serial_DataReceived;
            serialPort7.DataReceived += Serial_DataReceived;
        }

        private void CloseAllPorts()
        {
            TryClose(serialPort1);
            TryClose(serialPort2);
            TryClose(serialPort3);
            TryClose(serialPort4);
            TryClose(serialPort5);
            TryClose(serialPort6);
            TryClose(serialPort7);
        }

        private void OpenAllPorts()
        {
            TryOpen(serialPort1);
            TryOpen(serialPort2);
            TryOpen(serialPort3);
            TryOpen(serialPort4);
            TryOpen(serialPort5);
            TryOpen(serialPort6);
            TryOpen(serialPort7); // Modbus RTU
        }

        private void TryClose(SerialPort port)
        {
            if (port == null) return;

            try
            {
                if (port.IsOpen)
                {
                    port.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{port.PortName} 关闭失败: {ex.Message}");
            }
        }

        private void TryOpen(SerialPort port)
        {
            if (port == null) return;

            try
            {
                if (!port.IsOpen)
                {
                    port.Open();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{port.PortName} 打开失败: {ex.Message}");
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            initComList();
            this.switchMES.Click += SwitchMES_Click;
            this.RadioDebugMode.Click += RadioDebugMode_Click;
            this.RadioBtnProductionMode.Click += RadioBtnProductionMode_Click;
            this.txtMasterInput.TextChanged += TxtMasterInput_TextChanged;
            this.txtMasterInput.KeyDown += TxtMasterInput_KeyDown;
            this.btnReSet.Click += BtnReSet_Click;
            this.Shown += FrmMaster_Shown;
        }

        private void FrmMaster_Shown(object sender, EventArgs e)
        {
            InitScanMaster();
            initTitlePanelColor();
            maxCount = IntegerUpDownChannels.Value;
        }

        private void BtnReSet_Click(object sender, EventArgs e)
        {
            snList.Clear();
            txtsn1.Text = string.Empty;
            txtsn2.Text = string.Empty;
            txtsn3.Text = string.Empty;
            txtsn4.Text = string.Empty;
            txtsn5.Text = string.Empty;
            txtsn6.Text = string.Empty;
            InitScanMaster();
            this.ShowSuccessTip("队列已经清空");
        }

        private void TxtMasterInput_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                string barcode = txtMasterInput.Text.Trim();

                if (!string.IsNullOrEmpty(barcode))
                {
                    if (flagProductionModel)
                    {
                        var message = string.Empty;
                        var flag = MES_Service.CheckSerialNumber(barcode, ref message);
                        if (!flag)
                        {
                            this.ShowErrorDialog(message);
                            this.ShowErrorNotifier(message);
                            return;
                        }
                    }
                    snList.Add(new ScanModel { serialNumber = barcode, model = flagProductionModel });
                    // 👉 按当前数量安全赋值
                    if (snList.Count > 0) txtsn1.Text = snList[0].serialNumber;
                    if (snList.Count > 1) txtsn2.Text = snList[1].serialNumber;
                    if (snList.Count > 2) txtsn3.Text = snList[2].serialNumber;
                    if (snList.Count > 3) txtsn4.Text = snList[3].serialNumber;
                    if (snList.Count > 4) txtsn5.Text = snList[4].serialNumber;
                    if (snList.Count > 5) txtsn6.Text = snList[5].serialNumber;

                    // 👉 满6个再处理
                    if (snList.Count == maxCount)
                    {
                        // this.ShowErrorNotifier("已扫满6个");
                        // 这里可以做提交逻辑
                        this.ShowWaitForm();
                        // 如果要继续扫下一批，记得清空
                        this.HideWaitForm();
                        snList.Clear();
                    }

                    txtMasterInput.Clear();
                }

                e.SuppressKeyPress = true;
            }
        }

        private void TxtMasterInput_TextChanged(object sender, EventArgs e)
        {
            int length = txtMasterInput.Text.Length;
            lineLen.Text = $"长度: {length}";
        }

        /// <summary>
        /// Handles the click event for enabling or disabling radio debug mode.
        /// </summary>
        /// <param name="sender">Represents the source of the click event.</param>
        /// <param name="e">Contains the event data associated with the click action.</param>
        private void RadioDebugMode_Click(object sender, EventArgs e)
        {
            flagProductionModel = false;
        }

        private void SwitchMES_Click(object sender, EventArgs e)
        {
            //ERR.Visible = true;
            if (switchMES.Active)
            {
                var message = MES_Service.CheckLib();
                if (string.IsNullOrEmpty(message))
                {
                    //ShowSystemLogs("MES", "MesConnect");
                    MES_Service.MesConnect();
                }
                else
                {
                    this.ShowErrorNotifier($"MES连接失败 {message}");
                    switchMES.Active = false;
                    return;
                }
            }
            else
            {
                //ShowSystemLogs("MES", "MesDisConnect");
                MES_Service.MesDisConnect();
                //切换Debug模式
                // RadioDebugMode.Checked = true;
            }
        }

        /// <summary>
        /// Handles the click event for a radio button that toggles production mode.
        /// </summary>
        /// <param name="sender">Represents the source of the click event.</param>
        /// <param name="e">Contains the event data associated with the click action.</param>
        private void RadioBtnProductionMode_Click(object sender, EventArgs e)
        {
            if (RadioBtnProductionMode.Checked)
            {
                //if (string.IsNullOrEmpty(txtModelName.Text))
                //{
                //    RadioDebugMode.Checked = true;
                //    this.ShowErrorDialog("没有选择配置文件");
                //    return;
                //}

                if (string.IsNullOrEmpty(txtEmp.Text))
                {
                    RadioDebugMode.Checked = true;
                    this.ShowErrorDialog("没有输入人员工号");

                    return;
                }
                string message = string.Empty;
                bool checkFlag = MES_Service.CommandCheckUser(txtEmp.Text, ref message);
                if (!checkFlag)
                {
                    Sys_User.username = txtEmp.Text;
                    RadioDebugMode.Checked = true;

                    this.ShowErrorDialog($"{message}");
                    return;
                }
                Sys_User.username = txtEmp.Text;
                txtEmp.Enabled = false;
                txtMasterInput.SelectAll();
                txtMasterInput.Focus();
                this.ShowSuccessTip($"登录成功 {txtEmp.Text}");
            }
        }

        /// <summary>
        /// 设置LED颜色
        /// </summary>
        private void SetLedColor(System.IO.Ports.SerialPort port, dynamic led)
        {
            if (port == null)
            {
                led.Color = Color.Black;
            }
            else if (port.IsOpen)
            {
                led.Color = Color.Green;
            }
            else
            {
                led.Color = Color.Yellow;
            }
        }

        /// <summary>
        /// 初始化LED颜色
        /// </summary>
        private void initLedColor()
        {
            SetLedColor(serialPort1, ledCom1);
            SetLedColor(serialPort2, ledCom2);
            SetLedColor(serialPort3, ledCom3);
            SetLedColor(serialPort4, ledCom4);
            SetLedColor(serialPort5, ledCom5);
            SetLedColor(serialPort6, ledCom6);
            SetLedColor(serialPort7, ledCom7);
        }

        /// <summary>
        /// 初始化标题颜色
        /// </summary>
        public void initTitlePanelColor()
        {
            titlePanel1.TitleColor = Color.Gray;
            titlePanel2.TitleColor = Color.Gray;
            titlePanel3.TitleColor = Color.Gray;
            titlePanel4.TitleColor = Color.Gray;
            titlePanel5.TitleColor = Color.Gray;
            titlePanel6.TitleColor = Color.Gray;
            titlePanel8.TitleColor = Color.YellowGreen;
            //使用的状态是ForestGreen
            //待使用的状态是 yellowGreen
            if (IntegerUpDownChannels.Value > 0) { titlePanel1.TitleColor = Color.YellowGreen; }
            if (IntegerUpDownChannels.Value > 1) { titlePanel2.TitleColor = Color.YellowGreen; }
            if (IntegerUpDownChannels.Value > 2) { titlePanel3.TitleColor = Color.YellowGreen; }
            if (IntegerUpDownChannels.Value > 3) { titlePanel4.TitleColor = Color.YellowGreen; }
            if (IntegerUpDownChannels.Value > 4) { titlePanel5.TitleColor = Color.YellowGreen; }
            if (IntegerUpDownChannels.Value > 5) { titlePanel6.TitleColor = Color.YellowGreen; }
        }

        private void SetLedByIndex(int index, Color color)
        {
            switch (index)
            {
                case 0: ledCom1.Color = color; break;
                case 1: ledCom2.Color = color; break;
                case 2: ledCom3.Color = color; break;
                case 3: ledCom4.Color = color; break;
                case 4: ledCom5.Color = color; break;
                case 5: ledCom6.Color = color; break;
                case 6: ledCom7.Color = color; break;
            }
        }

        /// <summary>
        /// Initializes the LED indicators to reflect the current status of available serial ports (COM1–COM7).
        /// </summary>
        /// <remarks>This method updates the color of each corresponding LED indicator to green if the associated COM port
        /// is detected on the system. Call this method to synchronize the UI with the current set of available serial
        /// ports.</remarks>
        private void initComList()
        {
            initLedColor();

            for (int i = 1; i <= 7; i++)
            {
                string com = $"COM{i}";

                if (IsComExist(com))
                {
                    SetLedByIndex(i - 1, Color.Green);
                }
            }
        }

        /// <summary>
        /// 初始化
        /// </summary>
        private void InitScanMaster()
        {
            this.txtMasterInput.Text = string.Empty;
            this.txtMasterInput.Focus();
        }

        /// <summary>
        /// Determines whether the specified serial port name exists on the system.
        /// </summary>
        /// <param name="comName">The name of the serial port to check. This value is compared against the list of available serial ports.
        /// Cannot be null.</param>
        /// <returns>true if the specified serial port name exists; otherwise, false.</returns>
        private bool IsComExist(string comName)
        {
            return SerialPort.GetPortNames().Contains(comName);
        }

        /// <summary>
        /// Handles the DataReceived event of the serial port, which occurs when data is received through the port.
        /// </summary>
        /// <remarks>This method is typically used to process incoming serial data asynchronously. Ensure
        /// that any UI updates are marshaled to the UI thread, as this event is raised on a non-UI thread.</remarks>
        /// <param name="sender">The source of the event, typically the SerialPort instance that received data.</param>
        /// <param name="e">A SerialDataReceivedEventArgs object that contains the event data.</param>
        private void Serial_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            var port = sender as SerialPort;
            if (port == null) return;

            try
            {
                int len = port.BytesToRead;
                byte[] buffer = new byte[len];
                port.Read(buffer, 0, len);

                int index = GetPortIndex(port);

                // COM7 = Modbus
                if (port == serialPort7)
                {
                    // HandleModbus(buffer);
                }
                else
                {
                    HandleNormalSerial(index, buffer);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private int GetPortIndex(SerialPort port)
        {
            if (port == serialPort1) return 0;
            if (port == serialPort2) return 1;
            if (port == serialPort3) return 2;
            if (port == serialPort4) return 3;
            if (port == serialPort5) return 4;
            if (port == serialPort6) return 5;
            if (port == serialPort7) return 6;

            return -1;
        }

        private void HandleNormalSerial(int index, byte[] data)
        {
            string text = Encoding.ASCII.GetString(data);

            this.Invoke(new Action(() =>
            {
                switch (index)
                {
                    case 0: uiListBox1.Items.Add(text); break;
                    case 1: uiListBox2.Items.Add(text); break;
                    case 2: uiListBox3.Items.Add(text); break;
                    case 3: uiListBox4.Items.Add(text); break;
                    case 4: uiListBox5.Items.Add(text); break;
                    case 5: uiListBox6.Items.Add(text); break;
                }
            }));
        }
    }
}