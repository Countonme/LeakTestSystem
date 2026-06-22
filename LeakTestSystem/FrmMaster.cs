using Controller;
using LeakTestSystem.Controller;
using LeakTestSystem.Model;
using LeakTestSystem.Services;
using LeakTestSystem.Services.MES;
using Newtonsoft.Json;
using SNetLogs;
using Sunny.UI;
using Sunny.UI.Win32;
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
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
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
        private readonly string PASS = "PASS", FAIL = "FAIL";
        private readonly string OK = "OK", NG = "NG";
        private readonly string Testing = "Testing ...";
        private readonly StringBuilder _serialBuffer = new StringBuilder();
        private readonly object _lock = new object();
        private readonly object _logLock = new object();
        private readonly Dictionary<string, StringBuilder> _buffers = new Dictionary<string, StringBuilder>();
        private readonly UITextBox[] snTextBoxes;
        private readonly UILedDisplay[] _uiLedDisplaysArry;
        private readonly UIListBox[] _uiListBoxesArray;
        private SettingConfig _config = new SettingConfig();
        private FrmSN _frmSn;
        private readonly Dictionary<object, Color> _itemColorMap = new Dictionary<object, Color>();
        private string uuid;

        /// <summary>
        /// 测试记录
        /// </summary>
        private List<TestResult> testResults = new List<TestResult>();

        public FrmMaster()
        {
            InitializeComponent();
            this.Load += Form1_Load;
            this.switchMCUMaster.Click += SwitchMCUMaster_Click;
            this.Text += $"->(Version:{System.Windows.Forms.Application.ProductVersion})";
            this.FormClosing += FrmMaster_FormClosing;
            snTextBoxes = new[] { txtsn1, txtsn2, txtsn3, txtsn4, txtsn5, txtsn6 };
            _uiLedDisplaysArry = new[] { uiLedDisplay1, uiLedDisplay2, uiLedDisplay3, uiLedDisplay4, uiLedDisplay5, uiLedDisplay6 };
            _uiListBoxesArray = new[] { uiListBox1, uiListBox2, uiListBox3, uiListBox4, uiListBox5, uiListBox6 };
            foreach (var lb in _uiListBoxesArray)
            {
                //lb.DrawMode = DrawMode.OwnerDrawFixed;
                lb.DrawItem += ListBox_DrawItem;
            }
            //this.Shown += FrmMaster_Shown1;
        }

        private void FrmMaster_Shown1(object sender, EventArgs e)
        {
            //StartSN();
            var data = "<04>:-0.483 PSI:(OK):-0.1521 sccm";

            //HandleNormalSerial("COM1", data);
            var data2 = "<04>:-0.483 PSI:(AL):-0.1521 sccm";
            HandleNormalSerial("COM1", data);
            HandleNormalSerial("COM3", data);
            HandleNormalSerial("COM2", data2);
            HandleNormalSerial("COM1", data);
            HandleNormalSerial("COM2", data);
            HandleNormalSerial("COM3", data);
            HandleNormalSerial("COM2", data2);
            HandleNormalSerial("COM2", data);
            HandleNormalSerial("COM2", data2);
        }

        private void FrmMaster_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.ShowAskDialog("您确定退出窗体吗", true);
            {
                MES_Service.MesDisConnect();
                StopTesting();
                CloseAllPorts();
                Environment.Exit(0);
            }
        }

        /// <summary>
        /// 切换 MCU 主控的串口连接：启用时验证 COM1~COM7、初始化并打开所有串口、在 COM7 上初始化 Modbus 并在就绪后异步启动序列号检测；禁用时关闭所有串口。
        /// </summary>
        /// <remarks>启用流程按序执行：检查必需串口、初始化端口、打开端口并在 COM7 上创建 ModbusIoController。就绪后通过后台任务延迟调用
        /// StartSN。发生异常时复位开关状态、记录日志并显示错误通知。</remarks>
        /// <param name="sender">触发事件的发送者。</param>
        /// <param name="e">事件参数。</param>
        private void SwitchMCUMaster_Click(object sender, EventArgs e)
        {
            try
            {
                if (switchMCUMaster.Active)
                {
                    // 1. 判断是否存在 COM List 中启用的串口，如果缺失则提示并退出
                    if (!HasRequiredComPorts())
                    {
                        var missingPorts = GetMissingComPorts();
                        this.ShowErrorNotifier($"未检测到以下必要的串口: {string.Join(", ", missingPorts)}");
                        switchMCUMaster.Active = false;
                        return;
                    }
                    // 2. 初始化
                    InitAllPorts();

                    // 3. 打开所有串口
                    OpenAllPorts();

                    // 4. 重点：COM7 初始化 Modbus
                    modbusIo = new ModbusIoController(serialPort7);

                    this.ShowSuccessNotifier("所有串口已打开，Modbus已就绪 ...");
                    // ⭐ 关键修复
                    Task.Run(() =>
                    {
                        Thread.Sleep(3000);

                        this.BeginInvoke(new Action(() =>
                        {
                            StartSN();
                        }));
                    });
                }
                else
                {
                    CloseAllPorts();
                    this.ShowSuccessNotifier("所有串口已关闭");
                }
            }
            catch (Exception ex)
            {
                switchMCUMaster.Active = false;
                ShowLogs(ex.Message, Color.Red);
                this.ShowErrorNotifier(ex.Message);
            }
        }

        /// <summary>
        /// 判断系统是否有要使用的串口，至少要包含配方里启用的串口，否则无法正常使用
        /// </summary>
        /// <returns></returns>
        private bool HasRequiredComPorts()
        {
            return GetMissingComPorts().Count == 0;
        }

        private List<string> GetMissingComPorts()
        {
            var systemPorts = SerialPort.GetPortNames()
                .Select(p => p.Trim().ToUpper())
                .ToHashSet();

            var requiredPorts = _config.GetEnableComList(_config)
                .Select(p => p.Trim().ToUpper())
                .ToList();

            var missing = requiredPorts
                .Where(p => !systemPorts.Contains(p))
                .ToList();

            return missing;
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
                serialPort1.DataReceived += Serial_DataReceived;
            }
            if (_config.channel2Status)
            {
                serialPort2 = new SerialPort(_config.channel2ComName, 9600);
                serialPort2.DataReceived += Serial_DataReceived;
            }
            if (_config.channel3Status)
            {
                serialPort3 = new SerialPort(_config.channel3ComName, 9600);
                serialPort3.DataReceived += Serial_DataReceived;
            }
            if (_config.channel4Status)
            {
                serialPort4 = new SerialPort(_config.channel4ComName, 9600);
                serialPort4.DataReceived += Serial_DataReceived;
            }
            if (_config.channel5Status)
            {
                serialPort5 = new SerialPort(_config.channel5ComName, 9600);
                serialPort5.DataReceived += Serial_DataReceived;
            }
            if (_config.channel6Status)
            {
                serialPort6 = new SerialPort(_config.channel6ComName, 9600);
                serialPort6.DataReceived += Serial_DataReceived;
            }
            // COM7 = Modbus RTU
            serialPort7 = new SerialPort(_config.masterComName, 115200);
            // 统一绑定事件（关键）

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
            if (_config.channel1Status)
            {
                TryOpen(serialPort1);
            }
            if (_config.channel2Status)
            {
                TryOpen(serialPort2);
            }
            if (_config.channel3Status)
            {
                TryOpen(serialPort3);
            }
            if (_config.channel4Status)
            {
                TryOpen(serialPort4);
            }
            if (_config.channel5Status)
            {
                TryOpen(serialPort5);
            }
            if (_config.channel6Status)
            {
                TryOpen(serialPort6);
            }

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
            this.openToolStripMenuItem.Click += OpenToolStripMenuItem_Click;
            this.saveToolStripMenuItem.Click += SaveToolStripMenuItem_Click;
            this.refreshToolStripMenuItem.Click += RefreshToolStripMenuItem_Click;
            this.reloadToolStripMenuItem.Click += ReloadToolStripMenuItem_Click;

            InitUIDisplay("N/A", Color.Black);
            //InitUIDisplay("N/A", Color.Blue);
            //  InitUIDisplay("N/A", Color.Green);
            //InitUIDisplay("N/A", Color.Red);
            LoadConfig();
        }

        /// <summary>
        /// 显示打开文件对话框以选择并加载一个 .json 配置文件。
        /// </summary>
        /// <remarks>若不存在，则在 Application.StartupPath 下创建 proList 目录；选中文件后将其路径赋给 proName.Text，调用
        /// reloadToolStripMenuItem.PerformClick() 刷新，并根据结果显示成功或错误提示。</remarks>
        /// <param name="sender">事件的发送者。</param>
        /// <param name="e">事件的参数。</param>
        private void OpenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string baseDir = Path.Combine(Application.StartupPath, "proList");

            if (!Directory.Exists(baseDir))
                Directory.CreateDirectory(baseDir);

            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.InitialDirectory = baseDir;
                dlg.Filter = "Config File (*.json)|*.json";
                dlg.Title = "请选择配置文件";

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        proName.Text = dlg.FileName;
                        // ✔ 刷新系统
                        reloadToolStripMenuItem.PerformClick();

                        this.ShowSuccessTip("加载成功");
                    }
                    catch (Exception ex)
                    {
                        this.ShowErrorDialog($"加载失败: {ex.Message}");
                    }
                }
            }
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
                mesNgLock = switchMesNgLock.Active,
                snLength = snCheckLen.Value,
                ngCode = txtNgCode.Text
            };

            string json = Newtonsoft.Json.JsonConvert.SerializeObject(conf, Newtonsoft.Json.Formatting.Indented);

            // ✔ 默认目录：程序根目录/proList
            string baseDir = Path.Combine(Application.StartupPath, "proList");

            if (!Directory.Exists(baseDir))
                Directory.CreateDirectory(baseDir);

            using (SaveFileDialog dlg = new SaveFileDialog())
            {
                dlg.InitialDirectory = baseDir;
                dlg.Filter = "Config File (*.json)|*.json";
                dlg.Title = "请选择保存配置文件位置";

                // 默认文件名
                dlg.FileName = "pro.json";

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    File.WriteAllText(dlg.FileName, json);

                    reloadToolStripMenuItem.PerformClick();
                    this.ShowSuccessTip("保存成功");
                }
            }
        }

        /// <summary>
        /// 加载配方
        /// </summary>
        private void LoadConfig()
        {
            if (File.Exists(proName.Text))
            {
                //关闭串口，避免占用无法读取配置
                switchMCUMaster.Active = false;
                CloseAllPorts();

                var configString = File.ReadAllText(proName.Text);
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
                snCheckLen.Value = _config.snLength;
                titlePanel8.Text = $"Master Channels ({maxCount})";
                txtNgCode.Text = _config.ngCode;
                this.ShowSuccessNotifier("配置已加载...");
            }
            else
            {
                this.ShowErrorDialog($"配置文件 {proName.Text} 不存在，请先保存配置");
            }
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
                if (barcode.Length == snLen.Left)
                {
                    this.HideWaitForm();
                }

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
                    var index = snList.Count - 1;
                    var indexCount = _config.GetEnableChannelIndex(_config, index) - 1;
                    WriteSnToTextBox(indexCount, snList[index].serialNumber);
                    //if (snList.Count > 0) txtsn1.Text = snList[0].serialNumber;
                    //if (snList.Count > 1) txtsn2.Text = snList[1].serialNumber;
                    //if (snList.Count > 2) txtsn3.Text = snList[2].serialNumber;
                    //if (snList.Count > 3) txtsn4.Text = snList[3].serialNumber;
                    //if (snList.Count > 4) txtsn5.Text = snList[4].serialNumber;
                    //if (snList.Count > 5) txtsn6.Text = snList[5].serialNumber;

                    // 👉 满6个再处理
                    if (snList.Count == maxCount)
                    {
                        if (serialPort7 != null && serialPort7.IsOpen)
                        {
                            resultList.Clear();
                            // this.ShowErrorNotifier("已扫满6个");

                            // 这里可以做提交逻辑
                            this.ShowWaitForm("启动测试...请等待...");
                            // 如果要继续扫下一批，记得清空

                            //Thread.Sleep(1000);
                            //this.HideWaitForm();

                            StartTesting();
                            ShowLogs("启动测试...请等待...", Color.Black);
                        }
                        else
                        {
                            this.ShowErrorNotifier($"请先打开MCU 串口 {_config.masterComName}");
                            ShowLogs($"请先打开MCU 串口 {_config.masterComName}", Color.Red);
                        }
                        snList.Clear();
                        txtMasterInput.Focus();
                        txtMasterInput.SelectAll();
                    }
                }

                e.SuppressKeyPress = true;
            }
        }

        private void StartTesting()
        {
            this.Style = UIStyle.Blue;
            InitUIDisplay("N/A", Color.Yellow);
            //开启继电器
            if (_config.channel1Status)
            {
                modbusIo.SetRelay(1, 0, false);
                Thread.Sleep(100);
                ShowLogs("初始化继电器 1", Color.Black);
                modbusIo.SetRelay(1, 0, true); // 举例：触发继电器1
                InitUIDisplay(Testing, 0, Color.Blue);
            }
            if (_config.channel2Status)
            {
                modbusIo.SetRelay(1, 2, false);
                Thread.Sleep(100);
                ShowLogs("初始化继电器 2", Color.Black);
                modbusIo.SetRelay(1, 2, true); // 举例：触发继电器1
                InitUIDisplay(Testing, 1, Color.Blue);
            }
            if (_config.channel3Status)
            {
                modbusIo.SetRelay(1, 4, false);
                Thread.Sleep(100);
                ShowLogs("初始化继电器 3", Color.Black);
                modbusIo.SetRelay(1, 4, true); // 举例：触发继电器1
                InitUIDisplay(Testing, 2, Color.Blue);
            }
            if (_config.channel4Status)
            {
                modbusIo.SetRelay(1, 6, false);
                Thread.Sleep(100);
                ShowLogs("初始化继电器 4", Color.Black);
                modbusIo.SetRelay(1, 6, true); // 举例：触发继电器1
                InitUIDisplay(Testing, 3, Color.Blue);
            }
            if (_config.channel5Status)
            {
                modbusIo.SetRelay(1, 8, false);
                Thread.Sleep(100);
                ShowLogs("初始化继电器 5", Color.Black);
                modbusIo.SetRelay(1, 8, true); // 举例：触发继电器1
                InitUIDisplay(Testing, 4, Color.Blue);
            }
            if (_config.channel6Status)
            {
                modbusIo.SetRelay(1, 10, false);
                Thread.Sleep(100);
                ShowLogs("初始化继电器 6", Color.Black);
                modbusIo.SetRelay(1, 10, true); // 举例：触发继电器1
                InitUIDisplay(Testing, 5, Color.Blue);
            }
        }

        /// <summary>
        /// 停止测试，关闭所有继电器，并隐藏等待界面
        /// </summary>
        private void StopTesting()
        {
            //this.HideWaitForm();
            //开启继电器
            if (_config.channel1Status)
            {
                modbusIo.SetRelay(1, 0, false);
                Thread.Sleep(100);
                ShowLogs("关闭继电器 1", Color.Black);
            }
            if (_config.channel2Status)
            {
                modbusIo.SetRelay(1, 2, false);
                Thread.Sleep(100);
                ShowLogs("关闭继电器 2", Color.Black);
            }
            if (_config.channel3Status)
            {
                modbusIo.SetRelay(1, 4, false);
                Thread.Sleep(100);
                ShowLogs("关闭继电器 3", Color.Black);
            }
            if (_config.channel4Status)
            {
                modbusIo.SetRelay(1, 6, false);
                Thread.Sleep(100);
                ShowLogs("关闭继电器 4", Color.Black);
            }
            if (_config.channel5Status)
            {
                modbusIo.SetRelay(1, 8, false);
                Thread.Sleep(100);
                ShowLogs("关闭继电器 5", Color.Black);
            }
            if (_config.channel6Status)
            {
                modbusIo.SetRelay(1, 10, false);
                Thread.Sleep(100);
                ShowLogs("关闭继电器 6", Color.Black);
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
        //private void Serial_DataReceived(object sender, SerialDataReceivedEventArgs e)
        //{
        //    var port = sender as SerialPort;
        //    if (port == null) return;

        //    try
        //    {
        //        //int index = GetPortIndex(port);

        //        // COM7 = Modbus
        //        if (port == serialPort7)
        //        {
        //            // HandleModbus(buffer);
        //            byte[] data = new byte[port.BytesToRead];
        //            port.Read(data, 0, data.Length);

        //            string hex = BitConverter.ToString(data).Replace("-", " ");
        //            ShowLogs($"MCU COM7 收到数据: {hex}", Color.Blue);
        //            // ShowLogs($"MCU COM7 收到数据: {port.ReadExisting()}", Color.Blue);
        //        }
        //        else
        //        {
        //            //modbusIo.SetRelay(1, 0, false); // 举例：关闭继电器1
        //            this.HideWaitForm();
        //            StopTesting();
        //            string data = port.ReadExisting();

        //            lock (_lock)
        //            {
        //                _serialBuffer.Append(data);

        //                // 判断是否收到完整包
        //                while (_serialBuffer.ToString().Contains("\r\n"))
        //                {
        //                    string allData = _serialBuffer.ToString();

        //                    int endIndex = allData.IndexOf("\r\n");

        //                    // 取一包完整数据
        //                    string oneMessage = allData.Substring(0, endIndex);

        //                    // 删除已处理数据
        //                    _serialBuffer.Remove(0, endIndex + 2);

        //                    // int index = GetPortIndex(port);

        //                    this.BeginInvoke(new Action(() =>
        //                    {
        //                        this.ShowInfoTip(oneMessage);
        //                    }));

        //                    HandleNormalSerial(port.PortName, oneMessage);
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine(ex.Message);
        //    }
        //}
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
                if (port != serialPort7)
                {
                    StopTesting();
                    byte[] data = new byte[port.BytesToRead];

                    int len = port.Read(data, 0, data.Length);

                    if (len <= 0)
                        return;

                    string text = Encoding.ASCII.GetString(data, 0, len);

                    lock (_lock)
                    {
                        if (!_buffers.ContainsKey(port.PortName))
                        {
                            _buffers[port.PortName] = new StringBuilder();
                        }

                        var buffer = _buffers[port.PortName];

                        buffer.Append(text);

                        while (true)
                        {
                            string all = buffer.ToString();

                            int idx = all.IndexOf("\r\n");

                            if (idx < 0)
                                break;

                            string msg = all.Substring(0, idx);

                            buffer.Remove(0, idx + 2);

                            HandleNormalSerial(port.PortName, msg);
                        }
                    }
                }
                else
                {
                    // HandleModbus(buffer);
                    byte[] data = new byte[port.BytesToRead];
                    port.Read(data, 0, data.Length);
                    string hex = BitConverter.ToString(data).Replace("-", " ");
                    //ShowLogs($"MCU 收到数据: {hex}", Color.Blue);
                    // ShowLogs($"MCU COM7 收到数据: {port.ReadExisting()}", Color.Blue);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                ShowLogs($"读取串口 {port.PortName} 数据失败: {ex.Message}", Color.Red);
            }
        }

        private void WriteSnToTextBox(int index, string text)
        {
            switch (index)
            {
                case 0: txtsn1.Text = (text); break;
                case 1: txtsn2.Text = (text); break;
                case 2: txtsn3.Text = (text); break;
                case 3: txtsn4.Text = (text); break;
                case 4: txtsn5.Text = (text); break;
                case 5: txtsn6.Text = (text); break;
            }
        }

        /// <summary>
        /// COM 数据处理
        /// </summary>
        /// <param name="comName"></param>
        /// <param name="data"></param>
        private void HandleNormalSerial(string comName, string data)
        {
            try
            {
                string text = data;

                var index = _config.GetChannelIndexByComName(_config, comName);
                var listBox = _uiListBoxesArray[index];
                //var listBox = _uiListBoxesArray[2];

                //Color color = data.IndexOf("(OK)") > -1 ? Color.Red : Color.LimeGreen;
                bool isOk = data.Contains("(OK)");

                Color color = isOk
                    ? Color.LimeGreen
                    : Color.Red;

                listBox.Items.Insert(0, new ListBoxItemModel
                {
                    Text = data,
                    Color = color
                });

                while (listBox.Items.Count > 6)
                {
                    listBox.Items.RemoveAt(listBox.Items.Count - 1);
                }

                ShowLogs(data, color);
                GetResult(index, data, comName);
            }
            catch (Exception ex)
            {
                this.ShowErrorNotifier($"HandleNormalSerial {ex.Message}");
            }
        }

        private void ListBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;

            ListBox listBox = sender as ListBox;

            var item = listBox.Items[e.Index] as ListBoxItemModel;

            if (item == null) return;

            e.DrawBackground();

            using (Brush brush = new SolidBrush(item.Color))
            {
                e.Graphics.DrawString(
                    item.Text,
                    e.Font,
                    brush,
                    e.Bounds);
            }

            e.DrawFocusRectangle();
        }

        //private void HandleNormalSerial(string comName, string data)
        //{
        //    string text = data;

        //    var index = _config.GetChannelIndexByComName(_config, comName);
        //    var listBox = _uiListBoxesArray[index];

        //    // 判断状态
        //    Color color = Color.Red;

        //    if (data.Contains("(OK)"))
        //    {
        //        color = Color.LimeGreen;

        //    }
        //    else
        //    {
        //        color = Color.Red;
        //    }

        //    // 插入到第一行
        //    listBox.Items.Insert(0, text);

        //    while (listBox.Items.Count > 9)
        //    {
        //        listBox.Items.RemoveAt(listBox.Items.Count - 1);
        //    }

        //    ShowLogs(data, color);

        //    GetResult(index, data, comName);
        //}
        //private void ShowLogs(string InfoLogs, Color color)
        //{
        //    // 1. 如果日志太长，清空内容（保留你原有的逻辑）
        //    if (uiRichTextBox1.Text.Length > 1024)
        //    {
        //        uiRichTextBox1.Clear();
        //    }

        //    // 2. 构建要显示的字符串
        //    string msg = $"{DateTime.Now:HH:mm:ss} - {InfoLogs}";

        //    // 3. 【关键步骤】记录追加前的光标位置（也就是新文本开始的位置）
        //    int startIndex = uiRichTextBox1.TextLength;

        //    // 4. 追加文本（暂时不换行，或者先追加，后面一起处理）
        //    uiRichTextBox1.AppendText(msg);

        //    // 5. 【关键步骤】确定选中的长度
        //    // 我们需要选中从 startIndex 开始，到文本结束的部分
        //    // 注意：这里我们不需要手动加 Environment.NewLine，因为 AppendText 后光标已经在末尾
        //    int length = uiRichTextBox1.TextLength - startIndex;

        //    // 6. 【核心】设置颜色
        //    uiRichTextBox1.Select(startIndex, length); // 选中刚写入的内容
        //    uiRichTextBox1.SelectionColor = color;     // 设置颜色
        //    uiRichTextBox1.SelectionBackColor = uiRichTextBox1.BackColor; // 确保背景色不变

        //    // 7. 【关键步骤】添加换行符（为了下一行日志能正常换行）
        //    // 注意：换行符不能带颜色，所以必须在设置完颜色后单独追加
        //    uiRichTextBox1.AppendText(Environment.NewLine);

        //    // 8. 取消选中状态（将光标移到最后，避免界面上有一块蓝/红背景选中效果）
        //    uiRichTextBox1.Select(uiRichTextBox1.TextLength, 0);

        //    // 9. 自动滚动到底部
        //    uiRichTextBox1.ScrollToCaret();

        //    // --- 下面是你的文件保存逻辑 ---
        //    try
        //    {
        //        string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Logs", DateTime.Now.ToString("yyyyMM"));
        //        if (!Directory.Exists(path))
        //        {
        //            Directory.CreateDirectory(path);
        //        }
        //        string logFilePath = Path.Combine(path, $"{DateTime.Now:yyyyMMdd}.log");
        //        // 写入文件时不需要颜色，直接写原始文本即可
        //        File.AppendAllText(logFilePath, msg + Environment.NewLine);
        //    }
        //    catch (Exception ex)
        //    {
        //        // 简单的异常处理，防止日志写入失败导致程序崩溃
        //        MessageBox.Show($"写入日志失败: {ex.Message}");
        //    }
        //}

        private void ShowLogs(string infoLogs, Color color)
        {
            try
            {
                string msg = $"{DateTime.Now:HH:mm:ss.fff} - {infoLogs}";

                // UI线程安全
                if (uiRichTextBox1.InvokeRequired)
                {
                    uiRichTextBox1.BeginInvoke(new Action(() =>
                    {
                        AppendLogToUi(msg, color);
                    }));
                }
                else
                {
                    AppendLogToUi(msg, color);
                }

                // 写日志文件
                Log.Info(msg);
            }
            catch
            {
                // 日志系统不要抛异常
            }
        }

        private void AppendLogToUi(string msg, Color color)
        {
            try
            {
                // 控制最大长度
                if (uiRichTextBox1.TextLength > 1024 * 500)
                {
                    uiRichTextBox1.Clear();
                }

                uiRichTextBox1.SuspendLayout();

                int start = uiRichTextBox1.TextLength;

                uiRichTextBox1.AppendText(msg);

                uiRichTextBox1.Select(start, msg.Length);

                uiRichTextBox1.SelectionColor = color;
                uiRichTextBox1.SelectionBackColor = uiRichTextBox1.BackColor;

                uiRichTextBox1.AppendText(Environment.NewLine);

                uiRichTextBox1.SelectionLength = 0;
                uiRichTextBox1.ScrollToCaret();

                uiRichTextBox1.ResumeLayout();
            }
            catch
            {
            }
        }

        //private List<string> GetSnFromUser()
        //{
        //    _frmSn = new FrmSN(
        //        _config.GetEnableChannelCount(_config),
        //        _config.snLength,
        //        switchMES.Active);

        //    _frmSn.ShowDialog(); // 等待用户输入

        //    return _frmSn.GetAllSN();
        //}

        private List<string> GetSnFromUser()
        {
            _frmSn = new FrmSN(
               _config.GetEnableChannelCount(_config),
               _config.snLength,
               switchMES.Active);

            _frmSn.ShowDialog();

            return _frmSn.GetAllSN();
        }

        private void StartSN()
        {
            var snLists = GetSnFromUser();

            if (!_frmSn.IsAllOk())
            {
                this.ShowErrorNotifier("SN未全部校验通过");
                return;
            }

            if (snLists.Any(string.IsNullOrWhiteSpace))
            {
                this.ShowErrorNotifier("存在空SN");
                return;
            }

            snLists.ForEach(s =>
            {
                this.snList.Add(new ScanModel
                {
                    serialNumber = s,
                    model = flagProductionModel
                });
                // 👉 按当前数量安全赋值
                var index = snList.Count - 1;
                var indexCount = _config.GetEnableChannelIndex(_config, index) - 1;
                WriteSnToTextBox(indexCount, this.snList[index].serialNumber);
            });

            if (this.snList.Count == maxCount)
            {
                if (serialPort7 == null || !serialPort7.IsOpen)
                {
                    this.ShowErrorNotifier($"请先打开MCU串口 {_config.masterComName}");
                    return;
                }
                this.snList.Clear();
                resultList.Clear();
                var result = MessageBox.Show(
                    "队列条码已经准备好，您确定要启动测试吗？",
                    "提示",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );

                if (result == DialogResult.Yes)
                {
                    uuid = Guid.NewGuid().ToString("N");
                    //this.ShowWaitForm("启动测试...请等待...");
                    ShowLogs("启动测试...请等待...", Color.Black);
                    StartTesting();
                    // ⭐ 关键修复
                    Task.Run(() =>
                    {
                        Thread.Sleep(3000);

                        this.BeginInvoke(new Action(() =>
                        {
                            StartSN();
                        }));
                    });
                }
            }
        }

        private List<TestResult> resultList = new List<TestResult>();

        private void SettingColor(TestResult result, string status)
        {
            var index = _config.GetChannelIndexByComName(_config, result.comName);

            ShowLogs($"ComName:{result.comName} ComIndex{index} TestCount:{resultList.Count}  TotalCount:{_config.GetEnableChannelCount(_config)}", Color.DarkGreen);

            if (status == PASS)
            {
                InitUIDisplay(PASS, index, Color.Green);
            }
            else
            {
                InitUIDisplay(FAIL, index, Color.Red);
            }
        }

        private void AddTestHistory(string sn, TestResult result, string status, string message, string uuid)
        {
            var index = _config.GetChannelIndexByComName(_config, result.comName);

            ShowLogs($"ComName:{result.comName} ComIndex{index} TestCount:{resultList.Count}  TotalCount:{_config.GetEnableChannelCount(_config)}", Color.DarkGreen);
            result.testResult = status;
            result.serialNumber = sn;
            lock (resultList)
            {
                resultList.Add(result);
            }
            //多线程传
            Task.Run(() =>
            {
                UploadMesSystem(uuid, index, result, status, message);
            });

            ShowLogs($"TestCount:{resultList.Count}  TotalCount:{_config.GetEnableChannelCount(_config)}", Color.DarkGreen);

            if (resultList.Count == _config.GetEnableChannelCount(_config))
            {
                bool hasNg = resultList.Any(e => e.testResult == "NG");

                if (hasNg)
                {
                    //new pageResult("FAIL").ShowDialog();
                    this.Style = UIStyle.Red;
                }
                else
                {
                    //new pageResult("PASS").ShowDialog();
                }
                ShowLogs($"[MES全部完成] UUID:{uuid} {resultList.Count}:{maxCount}", Color.Green);

                this.Style = UIStyle.Green;
                resultList.Clear();
            }
        }

        private void UploadMesSystem(string uuid, int index, TestResult result, string status, string message)
        {
            var snresult = result.testResult == "OK" ? PASS : FAIL;
            if (switchMES.Active)
            {
                try
                {
                    ShowLogs($"[MES开始] UUID:{uuid} 总数:{maxCount}", Color.DarkBlue);

                    WriteExcel(uuid, result.serialNumber, result, snresult, message);
                    ShowLogs($"[MES处理开始] UUID:{uuid} {resultList.Count}/{maxCount} SN:{result.serialNumber} TestResult:{result.testResult}", Color.DarkBlue);

                    var listItem = new List<string>
                            {
                                $"PressureValue:{result.PressureValue}:{snresult}",
                                $"LeakValue:{result.LeakValue}:{snresult}"
                            };

                    // =========================
                    // NG也上传
                    // =========================
                    if (switchMesNgLock.Active)
                    {
                        ShowLogs($"[上传测试记录] UUID:{uuid} {resultList.Count}/{maxCount} SN:{result.serialNumber}", Color.DarkBlue);

                        if (!MES_Service.UploadTestRecords(result.serialNumber, listItem, ref message))
                        {
                            ShowLogs($"[上传测试记录失败] UUID:{uuid} {resultList.Count}/{maxCount} SN:{result.serialNumber} Message:{message}", Color.Red);
                            this.Style = UIStyle.Red;
                            InitUIDisplay("上传记录失败", index, Color.Red);
                            this.ShowErrorDialog(message);
                            return;
                        }

                        ShowLogs($"[上传测试记录成功] UUID:{uuid} {resultList.Count}/{maxCount} SN:{result.serialNumber}", Color.Green);

                        ShowLogs($"[开始过站] UUID:{uuid} {resultList.Count}/{maxCount} SN:{result.serialNumber}", Color.DarkBlue);

                        if (!MES_Service.SerialNumberCorssingStationFail(result.serialNumber, _config.ngCode, ref message))
                        {
                            this.Style = UIStyle.Red;
                            InitUIDisplay("锁定失败", index, Color.Red);
                            ShowLogs($"[过站失败] UUID:{uuid} {resultList.Count}/{maxCount} SN:{result.serialNumber} Message:{message}", Color.Red);
                            this.ShowErrorDialog($"[过站失败] UUID:{uuid} {resultList.Count}/{maxCount} SN:{result.serialNumber} Message:{message}");

                            return;
                        }
                        InitUIDisplay("不良品锁定" + FAIL, index, Color.Red);
                        ShowLogs($"[不良品 过站成功] UUID:{uuid} {resultList.Count}/{maxCount} SN:{result.serialNumber}", Color.Green);
                    }
                    else
                    {
                        // =========================
                        // 仅OK上传
                        // =========================
                        if (result.testResult == "OK")
                        {
                            ShowLogs($"[OK产品上传MES] UUID:{uuid} {resultList.Count}/{maxCount} SN:{result.serialNumber}", Color.DarkBlue);

                            ShowLogs($"[上传测试记录] UUID:{uuid} {resultList.Count}/{maxCount}  SN:{result.serialNumber}", Color.DarkBlue);

                            if (!MES_Service.UploadTestRecords(result.serialNumber, listItem, ref message))
                            {
                                this.Style = UIStyle.Red;
                                ShowLogs($"[上传测试记录失败] UUID:{uuid} {resultList.Count}/{maxCount}  SN:{result.serialNumber} Message:{message}", Color.Red);
                                InitUIDisplay("上传记录失败", index, Color.Red);
                                this.ShowErrorDialog(message);
                                return;
                            }

                            ShowLogs($"[上传测试记录成功] UUID:{uuid} {resultList.Count}/{maxCount}  SN:{result.serialNumber}", Color.Green);

                            ShowLogs($"[开始过站] UUID:{uuid} {resultList.Count}/{maxCount}  SN:{result.serialNumber}", Color.DarkBlue);

                            if (!MES_Service.SerialNumberCorssingStationPass(result.serialNumber, ref message))
                            {
                                this.Style = UIStyle.Red;
                                InitUIDisplay("过站失败", index, Color.Red);
                                ShowLogs($"[过站失败] UUID:{uuid} {resultList.Count}/{maxCount} SN:{result.serialNumber} Message:{message}", Color.Red);
                                this.ShowErrorDialog($"[过站失败] UUID:{uuid} {resultList.Count}/{maxCount}  SN:{result.serialNumber} Message:{message}");
                                return;
                            }
                            InitUIDisplay("过站成功" + PASS, index, Color.Green);
                            ShowLogs($"[过站成功] UUID:{uuid} {resultList.Count}/{maxCount}  SN:{result.serialNumber}", Color.Green);
                        }
                        else
                        {
                            InitUIDisplay("NG跳过MES", index, Color.IndianRed);
                            ShowLogs($"[NG跳过MES] UUID:{uuid} {resultList.Count}/{maxCount}  SN:{result.serialNumber}", Color.Orange);
                        }
                    }

                    ShowLogs($"[MES处理完成] UUID:{uuid} {resultList.Count}/{maxCount}  SN:{result.serialNumber}", Color.Green);
                }
                catch (Exception ex)
                {
                    this.Style = UIStyle.Red;
                    ShowLogs($"[MES异常] UUID:{uuid} {resultList.Count}/{maxCount}  SN:{result.serialNumber} Exception:{ex}", Color.Red);

                    this.ShowErrorDialog(ex.Message);
                    // return;
                }
                // Thread.Sleep(20);
            }
            else
            {
                if (result.testResult == "OK")
                {
                    InitUIDisplay("跳过MES PASS", index, Color.Green);
                }
                else
                {
                    InitUIDisplay("不良品", index, Color.Red);
                }

                WriteExcel(uuid, result.serialNumber, result, snresult, message);
            }
        }

        /// <summary>
        /// 解析测试结果并显示日志
        /// </summary>
        /// <param name="index">索引，用于获取对应的SN</param>
        /// <param name="data">测试结果数据</param>
        /// <summary>
        /// 解析测试结果并显示日志
        /// </summary>
        private void GetResult(int index, string data, string comName)
        {
            try
            {
                string sn = GetSN(index);

                if (!TestResult.TryParse(data, out TestResult result))
                {
                    ShowLogs($"数据格式错误 SN:{sn} Data:{data}", Color.Red);
                    result.serialNumber = sn;
                    result.comName = comName;
                    result.channelsName = $"CH{index + 1}";
                    //lock (resultList)
                    //{
                    //    resultList.Add(result);
                    //}
                    AddTestHistory(sn, result, "NG", "数据格式错误", uuid);
                    return;
                }

                string msg;
                result.serialNumber = sn;
                result.comName = comName;
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

                        AddTestHistory(sn, result, "OK", msg, uuid);

                        ShowLogs(msg, Color.Green);
                        break;

                    case "AL":

                        msg =
                            $"测试结果:SN:{sn} " +
                            $"报警:{result.alarmMessage} " +
                            $"通道:{result.channelsName}";

                        AddTestHistory(sn, result, "NG", msg, uuid);
                        ShowLogs(msg, Color.Red);
                        break;

                    case "TD":

                        msg =
                            $"测试结果:SN:{sn} " +
                            $"正常值泄漏NG " +
                            $"通道:{result.channelsName} " +
                            $"压力:{result.PressureValue} " +
                            $"泄漏:{result.LeakValue}";

                        AddTestHistory(sn, result, "NG", msg, uuid);
                        ShowLogs(msg, Color.Orange);
                        break;

                    case "RD":

                        msg =
                            $"测试结果:SN:{sn} " +
                            $"负值泄漏NG " +
                            $"通道:{result.channelsName} " +
                            $"压力:{result.PressureValue} " +
                            $"泄漏:{result.LeakValue}";

                        AddTestHistory(sn, result, "NG", msg, uuid);
                        ShowLogs(msg, Color.Orange);
                        break;

                    default:

                        msg =
                            $"测试结果:SN:{sn} " +
                            $"未知结果:{result.testResult} " +
                            $"原始数据:{data}";
                        AddTestHistory(sn, result, "NG", msg, uuid);
                        ShowLogs(msg, Color.Gray);
                        break;
                }
            }
            catch (Exception ex)
            {
                ShowLogs($"处理测试结果时发生异常: {ex.Message}", Color.Red);
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

        /// <summary>
        /// 初始化状态
        /// </summary>
        private void InitUIDisplay(string value)
        {
            //var value =String.Empty ;
            _uiLedDisplaysArry.ForEach(e =>
            {
                e.Text = value;
                e.ForeColor = Color.White;
                e.BackColor = Color.Black;
            });
        }

        /// <summary>
        /// 初始化 LED 显示状态
        /// </summary>
        /// <param name="value"></param>
        /// <param name="index"></param>
        /// <param name="color"></param>
        private void InitUIDisplay(string value, int index, Color color)
        {
            var ctrl = _uiLedDisplaysArry[index];
            if (ctrl.InvokeRequired)
            {
                ctrl.BeginInvoke(new Action(() =>
                {
                    ctrl.Text = value;
                    ctrl.LedBackColor = color;
                }));
            }
            else
            {
                ctrl.Text = value;
                ctrl.LedBackColor = color;
            }
        }

        /// <summary>
        /// 初始化状态
        /// </summary>
        private void InitUIDisplay(string value, Color color)
        {
            //var value =String.Empty ;
            _uiLedDisplaysArry.ForEach(e =>
            {
                e.BeginInvoke(new Action(() =>
                {
                    e.Text = value;
                    e.ForeColor = Color.White;
                    e.LedBackColor = color;
                }));
            });
        }

        private readonly object _excelLock = new object();

        private void WriteExcel(string uuid, string sn, TestResult result, string status, string message)
        {
            try
            {
                string dir = Path.Combine(Application.StartupPath, "TestExcel");
                if (!Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                string filePath = Path.Combine(dir, $"{DateTime.Now:yyyyMMdd}.xlsx");

                lock (_excelLock)
                {
                    using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(filePath)))
                    {
                        var sheet = package.Workbook.Worksheets.FirstOrDefault()
                                    ?? package.Workbook.Worksheets.Add("Test");

                        // =========================
                        // 1. 初始化表头
                        // =========================
                        if (sheet.Dimension == null)
                        {
                            string[] headers =
                            {
                       "UUID", "Time", "SN", "Channel", "Com",
                        "Pressure", "Leak", "Result", "Message"
                    };

                            for (int i = 0; i < headers.Length; i++)
                            {
                                var cell = sheet.Cells[1, i + 1];
                                cell.Value = headers[i];

                                // 表头样式
                                cell.Style.Font.Bold = true;
                                cell.Style.Font.Size = 11;
                                cell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                cell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(45, 85, 155));
                                cell.Style.Font.Color.SetColor(Color.White);
                            }

                            sheet.Row(1).Height = 22;

                            // 冻结首行
                            sheet.View.FreezePanes(2, 1);
                        }

                        int row = sheet.Dimension?.Rows + 1 ?? 2;
                        sheet.Cells[row, 1].Value = uuid;
                        sheet.Cells[row, 2].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        sheet.Cells[row, 3].Value = sn;
                        sheet.Cells[row, 4].Value = result.channelsName;
                        sheet.Cells[row, 5].Value = result.comName;
                        sheet.Cells[row, 6].Value = result.PressureValue;
                        sheet.Cells[row, 7].Value = result.LeakValue;
                        sheet.Cells[row, 8].Value = status;
                        sheet.Cells[row, 9].Value = message;

                        // =========================
                        // 2. 行样式（OK / NG）
                        // =========================
                        var rowRange = sheet.Cells[row, 1, row, 8];

                        rowRange.Style.HorizontalAlignment =
                            OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                        rowRange.Style.VerticalAlignment =
                            OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                        if (status == PASS)
                        {
                            rowRange.Style.Font.Color.SetColor(Color.DarkGreen);
                        }
                        else
                        {
                            rowRange.Style.Font.Color.SetColor(Color.Red);
                        }

                        // =========================
                        // 3. 自动列宽
                        // =========================
                        sheet.Cells[sheet.Dimension.Address].AutoFitColumns();

                        package.Save();
                    }
                }
            }
            catch (Exception ex)
            {
                ShowLogs($"Excel写入失败: {ex.Message}", Color.Red);
            }
        }
    }
}