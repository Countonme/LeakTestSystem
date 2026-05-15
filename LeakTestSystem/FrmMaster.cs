using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Controller;
using LeakTestSystem.Model;
using LeakTestSystem.Services;
using LeakTestSystem.Services.MES;
using Newtonsoft.Json;
using Sunny.UI;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Application = System.Windows.Forms.Application;

namespace LeakTestSystem
{
    public partial class FrmMaster : UIForm
    {
        public ModbusIoController modbusIo;
        private SerialPort serialPort1, serialPort2, serialPort3, serialPort4, serialPort5, serialPort6, serialPort7;
        private bool flagProductionModel;
        private List<ScanModel> snList = new List<ScanModel>();
        private int maxCount = 6;
        private readonly StringBuilder _serialBuffer = new StringBuilder();
        private readonly object _lock = new object();
        private readonly UITextBox[] snTextBoxes;
        private SettingConfig _config = new SettingConfig();
        /// <summary>
        /// 测试记录
        /// </summary>
        private List<TestResult> testResults = new List<TestResult>();

        public FrmMaster()
        {
            InitializeComponent();
            this.Load += Form1_Load;
            this.switchCom7.Click += SwitchCom7_Click;
            this.Text += $"->(Version:{System.Windows.Forms.Application.ProductVersion})";
            //this.FrmMaster_FormClosing += FrmMaster_FormClosing;
            snTextBoxes = new[] { txtsn1, txtsn2, txtsn3, txtsn4, txtsn5, txtsn6 };
        }

        private void FrmMaster_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.ShowAskDialog("您确定退出窗体吗", true);
            {
                MES_Service.MesDisConnect();
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

                this.ShowSuccessNotifier("所有串口已打开，Modbus已就绪");
            }
            else
            {
                CloseAllPorts();
                this.ShowSuccessNotifier("所有串口已关闭");
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
            if (string.IsNullOrEmpty(_config.masterComName))
            {
                this.ShowErrorNotifier(_config.masterComName);
                return;
            }
            if (_config.channel1Status)
            {
                serialPort1 = new SerialPort(_config.channel1ComName, 9600);
            }
            if (_config.channel2Status)
            {
                serialPort2 = new SerialPort(_config.channel2ComName, 9600);
            }
            if (_config.channel3Status)
            {
                serialPort3 = new SerialPort(_config.channel3ComName, 9600);
            }
            if (_config.channel4Status)
            {
                serialPort4 = new SerialPort(_config.channel4ComName, 9600);
            }
            if (_config.channel5Status)
            {
                serialPort5 = new SerialPort(_config.channel5ComName, 9600);
            }
            if (_config.channel6Status)
            {
                serialPort6 = new SerialPort(_config.channel6ComName, 9600);
            }
            // COM7 = Modbus RTU
            serialPort7 = new SerialPort(_config.masterComName, 115200);
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
                ShowLogs($"{port.PortName} 关闭失败: {ex.Message}", Color.Red);
                this.ShowErrorNotifier($"{port.PortName} 关闭失败: {ex.Message}");
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
                ShowLogs($"{port.PortName} 打开失败: {ex.Message}", Color.Red);
                this.ShowErrorNotifier($"{port.PortName} 打开失败: {ex.Message}");
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
            this.uiCheckBox1.Click += UiCheckBox_Click;
            this.uiCheckBox2.Click += UiCheckBox_Click;
            this.uiCheckBox3.Click += UiCheckBox_Click;
            this.uiCheckBox4.Click += UiCheckBox_Click;
            this.uiCheckBox5.Click += UiCheckBox_Click;
            this.uiCheckBox6.Click += UiCheckBox_Click;
            this.uiCheckBox7.Click += UiCheckBox_Click;
            this.uiCheckBox8.Click += UiCheckBox_Click;
            this.uiCheckBox9.Click += UiCheckBox_Click;
            this.uiCheckBox10.Click += UiCheckBox_Click;
            this.uiCheckBox11.Click += UiCheckBox_Click;
            this.uiCheckBox12.Click += UiCheckBox_Click;
            this.uiCheckBox13.Click += UiCheckBox_Click;
            this.uiCheckBox14.Click += UiCheckBox_Click;
            this.uiCheckBox15.Click += UiCheckBox_Click;
            this.uiCheckBox16.Click += UiCheckBox_Click;
            //var data = "<03>:10.95 kPa:(OK):0.1408 sccm";
            //var data = "<04>:-0.483 PSI:(OK):-0.1521 sccm";
            //GetResult(0, data);
            //var data1 = "<03>:(AL):SEALED PART VOL TOO SMALL";
            //GetResult(1, data1);
            this.saveToolStripMenuItem.Click += SaveToolStripMenuItem_Click;
            this.refreshToolStripMenuItem.Click += RefreshToolStripMenuItem_Click;
            this.reloadToolStripMenuItem.Click += ReloadToolStripMenuItem_Click;
            LoadConfig();
        }
        /// <summary>
        /// 重新加载配方
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ReloadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LoadConfig();
            initTitlePanelColor();
 
        }

        /// <summary>
        /// 重新加载COM列表，更新UI显示
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            reloadComList();
        }

        /// <summary>
        /// 保存配方
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <exception cref="NotImplementedException"></exception>
        private void SaveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var conf = new SettingConfig
            {
                masterComName = cobChannelMaster.Text,
                channel1ComName = cobChannel1.Text,
                channel2ComName = cobChannel2.Text,
                channel3ComName = cobChannel3.Text,
                channel4ComName = cobChannel4.Text,
                channel5ComName = cobChannel5.Text,
                channel6ComName = cobChannel6.Text,
                channel1Status = checkBoxChannel1.Checked,
                channel2Status = checkBoxChannel2.Checked,
                channel3Status = checkBoxChannel3.Checked,
                channel4Status = checkBoxChannel4.Checked,
                channel5Status = checkBoxChannel5.Checked,
                channel6Status = checkBoxChannel6.Checked,
                readTimeout = readTimeout.Value,
                mesNgLock = switchMesNgLock.Active
            };
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(conf, Newtonsoft.Json.Formatting.Indented);
            string path = Path.Combine(Application.StartupPath, "config.json");
            File.WriteAllText(path, json);
            reloadToolStripMenuItem.PerformClick();
            this.ShowSuccessTip("保存成功");
        }
        /// <summary>
        /// 加载配方
        /// </summary>
        private void LoadConfig()
        {
            if (File.Exists(Path.Combine(Application.StartupPath, "config.json")))
            {
                //关闭串口，避免占用无法读取配置
                switchCom7.Active = false;
                CloseAllPorts();

                var configString = File.ReadAllText(Path.Combine(Application.StartupPath, "config.json"));
                _config = JsonConvert.DeserializeObject<SettingConfig>(configString) ?? new SettingConfig();
                cobChannelMaster.Text = _config.masterComName;
                cobChannel1.Text = _config.channel1ComName;
                cobChannel2.Text = _config.channel2ComName;
                cobChannel3.Text = _config.channel3ComName;
                cobChannel4.Text = _config.channel4ComName;
                cobChannel5.Text = _config.channel5ComName;
                cobChannel6.Text = _config.channel6ComName;
                checkBoxChannel1.Checked = _config.channel1Status;
                checkBoxChannel2.Checked = _config.channel2Status;
                checkBoxChannel3.Checked = _config.channel3Status;
                checkBoxChannel4.Checked = _config.channel4Status;
                checkBoxChannel5.Checked = _config.channel5Status;
                checkBoxChannel6.Checked = _config.channel6Status;
                readTimeout.Value = _config.readTimeout;
                labCH1.Text = _config.channel1Status ? _config.channel1ComName : "未启用";
                labCH2.Text = _config.channel2Status ? _config.channel2ComName : "未启用";
                labCH3.Text = _config.channel3Status ? _config.channel3ComName : "未启用";
                labCH4.Text = _config.channel4Status ? _config.channel4ComName : "未启用";
                labCH5.Text = _config.channel5Status ? _config.channel5ComName : "未启用";
                labCH6.Text = _config.channel6Status ? _config.channel6ComName : "未启用";
                labMaster.Text = _config.masterComName;
                switchMesNgLock.Active = _config.mesNgLock;
                maxCount = _config.GetEnableChannelCount(_config);
                titlePanel8.Text = $"Master Channels ({maxCount})";
            }
            this.ShowSuccessNotifier("配置已加载...");
        }


        /// <summary>
        /// 重新加载串口
        /// </summary>
        private void reloadComList()
        {
            var com = SerialPort.GetPortNames();
            cobChannel1.Items.Clear();
            cobChannel2.Items.Clear();
            cobChannel3.Items.Clear();
            cobChannel4.Items.Clear();
            cobChannel5.Items.Clear();
            cobChannel6.Items.Clear();
            cobChannelMaster.Items.Clear();
            foreach (var comName in com)
            {
                cobChannel1.Items.Add(comName);
                cobChannel2.Items.Add(comName);
                cobChannel3.Items.Add(comName);
                cobChannel4.Items.Add(comName);
                cobChannel5.Items.Add(comName);
                cobChannel6.Items.Add(comName);
                cobChannelMaster.Items.Add(comName);
            }
            this.ShowSuccessNotifier($"COM 列表已刷新 共找到{com.Length}个串口");
        }

        private void UiCheckBox_Click(object sender, EventArgs e)
        {
            var checkBox = sender as UICheckBox;
            if (serialPort7 != null && serialPort7.IsOpen)
            {
                var cli = int.Parse(checkBox.Text) - 1;
                if (checkBox.Checked)
                {
                    this.ShowSuccessTip($"已选中 {cli}");
                    modbusIo.SetRelay(1, cli, true);
                }
                else
                {
                    this.ShowErrorTip($"已取消 {cli}");
                    modbusIo.SetRelay(1, cli, false);
                }
            }
            else
            {
                this.ShowErrorNotifier("请先打开COM7");
                ShowLogs("请先打开COM7", Color.Red);
            }
        }

        private void FrmMaster_Shown(object sender, EventArgs e)
        {
            InitScanMaster();
            initTitlePanelColor();
            uiGroupBox3.Height = this.Height - titlePanel1.Height - titlePanel1.Location.Y - 80;
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
                        if (serialPort7 != null && serialPort7.IsOpen)
                        {
                            // this.ShowErrorNotifier("已扫满6个");
                            modbusIo.SetRelay(1, 0, false);
                            ShowLogs("初始化继电器", Color.Black);
                            // 这里可以做提交逻辑
                            this.ShowWaitForm("启动测试...请等待...");
                            // 如果要继续扫下一批，记得清空

                            //Thread.Sleep(1000);
                            //this.HideWaitForm();
                            //开启继电器
                            modbusIo.SetRelay(1, 0, true); // 举例：触发继电器1
                            ShowLogs("启动测试...请等待...", Color.Black);
                            snList.Clear();
                        }
                        else
                        {
                            this.ShowErrorNotifier($"请先打开MCU 串口 {_config.masterComName}");
                            ShowLogs($"请先打开MCU 串口 {_config.masterComName}", Color.Red);
                        }
                        txtMasterInput.Clear();
                    }
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
                ShowLogs($"登录成功 {txtEmp.Text}", Color.Green);
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
            ConnectitonStatusV1.ValveColor = Color.Gray;
            ConnectitonStatusV2.ValveColor = Color.Gray;
            ConnectitonStatusV3.ValveColor = Color.Gray;
            ConnectitonStatusV4.ValveColor = Color.Gray;
            ConnectitonStatusV5.ValveColor = Color.Gray;
            ConnectitonStatusV6.ValveColor = Color.Gray;

            //使用的状态是ForestGreen
            //待使用的状态是 yellowGreen
            if (_config.channel1Status) { titlePanel1.TitleColor = Color.YellowGreen; ConnectitonStatusV1.ValveColor = Color.YellowGreen; }
            if (_config.channel2Status) { titlePanel2.TitleColor = Color.YellowGreen; ConnectitonStatusV2.ValveColor = Color.YellowGreen; }
            if (_config.channel3Status) { titlePanel3.TitleColor = Color.YellowGreen; ConnectitonStatusV3.ValveColor = Color.YellowGreen; }
            if (_config.channel4Status) { titlePanel4.TitleColor = Color.YellowGreen; ConnectitonStatusV4.ValveColor = Color.YellowGreen; }
            if (_config.channel5Status) { titlePanel5.TitleColor = Color.YellowGreen; ConnectitonStatusV5.ValveColor = Color.YellowGreen; }
            if (_config.channel6Status) { titlePanel6.TitleColor = Color.YellowGreen; ConnectitonStatusV6.ValveColor = Color.YellowGreen; }
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
                //int index = GetPortIndex(port);

                // COM7 = Modbus
                if (port == serialPort7)
                {
                    // HandleModbus(buffer);
                    byte[] data = new byte[port.BytesToRead];
                    port.Read(data, 0, data.Length);

                    string hex = BitConverter.ToString(data).Replace("-", " ");
                    ShowLogs($"MCU COM7 收到数据: {hex}", Color.Blue);
                    // ShowLogs($"MCU COM7 收到数据: {port.ReadExisting()}", Color.Blue);
                }
                else
                {
                    modbusIo.SetRelay(1, 0, false); // 举例：关闭继电器1
                    this.HideWaitForm();
                    string data = port.ReadExisting();

                    lock (_lock)
                    {
                        _serialBuffer.Append(data);

                        // 判断是否收到完整包
                        while (_serialBuffer.ToString().Contains("\r\n"))
                        {
                            string allData = _serialBuffer.ToString();

                            int endIndex = allData.IndexOf("\r\n");

                            // 取一包完整数据
                            string oneMessage = allData.Substring(0, endIndex);

                            // 删除已处理数据
                            _serialBuffer.Remove(0, endIndex + 2);

                            int index = GetPortIndex(port);

                            this.BeginInvoke(new Action(() =>
                            {
                                this.ShowInfoTip(oneMessage);
                            }));

                            HandleNormalSerial(index, oneMessage);
                        }
                    }
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

        private void HandleNormalSerial(int index, string data)
        {
            string text = (data);

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
            ShowLogs(data, Color.Black);
            //var data = "<03>:10.95 kPa:(OK):0.1408 sccm";
            GetResult(index, data);
        }

        private void ShowLogs(string InfoLogs, Color color)
        {
            // 1. 如果日志太长，清空内容（保留你原有的逻辑）
            if (uiRichTextBox1.Text.Length > 1024)
            {
                uiRichTextBox1.Clear();
            }

            // 2. 构建要显示的字符串
            string msg = $"{DateTime.Now:HH:mm:ss} - {InfoLogs}";

            // 3. 【关键步骤】记录追加前的光标位置（也就是新文本开始的位置）
            int startIndex = uiRichTextBox1.TextLength;

            // 4. 追加文本（暂时不换行，或者先追加，后面一起处理）
            uiRichTextBox1.AppendText(msg);

            // 5. 【关键步骤】确定选中的长度
            // 我们需要选中从 startIndex 开始，到文本结束的部分
            // 注意：这里我们不需要手动加 Environment.NewLine，因为 AppendText 后光标已经在末尾
            int length = uiRichTextBox1.TextLength - startIndex;

            // 6. 【核心】设置颜色
            uiRichTextBox1.Select(startIndex, length); // 选中刚写入的内容
            uiRichTextBox1.SelectionColor = color;     // 设置颜色
            uiRichTextBox1.SelectionBackColor = uiRichTextBox1.BackColor; // 确保背景色不变

            // 7. 【关键步骤】添加换行符（为了下一行日志能正常换行）
            // 注意：换行符不能带颜色，所以必须在设置完颜色后单独追加
            uiRichTextBox1.AppendText(Environment.NewLine);

            // 8. 取消选中状态（将光标移到最后，避免界面上有一块蓝/红背景选中效果）
            uiRichTextBox1.Select(uiRichTextBox1.TextLength, 0);

            // 9. 自动滚动到底部
            uiRichTextBox1.ScrollToCaret();

            // --- 下面是你的文件保存逻辑 ---
            try
            {
                string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Logs", DateTime.Now.ToString("yyyyMM"));
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                string logFilePath = Path.Combine(path, $"{DateTime.Now:yyyyMMdd}.log");
                // 写入文件时不需要颜色，直接写原始文本即可
                File.AppendAllText(logFilePath, msg + Environment.NewLine);
            }
            catch (Exception ex)
            {
                // 简单的异常处理，防止日志写入失败导致程序崩溃
                MessageBox.Show($"写入日志失败: {ex.Message}");
            }
        }

        private async Task GetWaitTestResult()
        {
        }

        /// <summary>
        /// 解析测试结果并显示日志
        /// </summary>
        /// <param name="index">索引，用于获取对应的SN</param>
        /// <param name="data">测试结果数据</param>
        /// <summary>
        /// 解析测试结果并显示日志
        /// </summary>
        private void GetResult(int index, string data)
        {
            string sn = GetSN(index);

            if (!TestResult.TryParse(data, out TestResult result))
            {
                ShowLogs($"数据格式错误 SN:{sn} Data:{data}", Color.Red);
                new pageResult("FAIL").ShowDialog();
                return;
            }

            string msg;
            result.channelsName = $"CH{index + 1}";
            switch (result.testResult)
            {
                case "OK":
                    msg =
                        $"测试结果:SN:{sn} " +
                        $"结果:{result.testResult} " +
                        $"通道:{result.channelsName} " +
                        $"压力:{result.PressureValue} " +
                        $"泄漏:{result.LeakValue}";

                    ShowLogs(msg, Color.Green);
                    new pageResult("PASS").ShowDialog();
                    break;

                case "AL":
                    msg =
                        $"测试结果:SN:{sn} " +
                        $"报警:{result.alarmMessage} " +
                        $"通道:{result.channelsName}";
                    new pageResult("FAIL").ShowDialog();
                    ShowLogs(msg, Color.Red);
                    break;

                case "TD":
                    msg =
                        $"测试结果:SN:{sn} " +
                        $"正常值泄漏NG " +
                        $"通道:{result.channelsName} " +
                        $"压力:{result.PressureValue} " +
                        $"泄漏:{result.LeakValue}";
                    new pageResult("FAIL").ShowDialog();
                    ShowLogs(msg, Color.Orange);
                    break;

                case "RD":
                    msg =
                        $"测试结果:SN:{sn} " +
                        $"负值泄漏NG " +
                        $"通道:{result.channelsName} " +
                        $"压力:{result.PressureValue} " +
                        $"泄漏:{result.LeakValue}";
                    new pageResult("FAIL").ShowDialog();
                    ShowLogs(msg, Color.Orange);
                    break;

                default:
                    msg =
                        $"测试结果:SN:{sn} " +
                        $"未知结果:{result.testResult} " +
                        $"原始数据:{data}";
                    new pageResult("FAIL").ShowDialog();
                    ShowLogs(msg, Color.Gray);
                    break;
            }
        }

        /// <summary>
        /// 获取SN，根据index返回对应的文本框内容
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        private string GetSN(int index)
        {
            if (index < 0 || index >= snTextBoxes.Length)
                return string.Empty;

            return snTextBoxes[index].Text.Trim();
        }
    }
}