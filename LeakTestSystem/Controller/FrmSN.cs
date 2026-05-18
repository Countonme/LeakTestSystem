using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using LeakTestSystem.Services.MES;
using Sunny.UI;

namespace LeakTestSystem.Controller
{
    public partial class FrmSN : UIForm
    {
        // 保存所有输入框
        private readonly List<UITextBox> _textBoxes = new List<UITextBox>();

        //private bool _allowClose = false;
        // 保存状态Label
        private readonly List<UILabel> _statusLabels = new List<UILabel>();

        private int _snLength = 0;
        private bool _mes = false;

        public FrmSN(int count, int snLength, bool mes)
        {
            InitializeComponent();
            var Mode = mes ? "MES" : "Debugger";
            this.Text += $" Model:{Mode} 标准SerialNumber长度:{snLength}  ";
            InitSNInput(count);
            this._snLength = snLength;
            this._mes = mes;
            this.Shown += OnShown;
        }

        private void OnShown(object sender, EventArgs e)
        {
            this.Top = 50;
        }

        //protected override void OnFormClosing(FormClosingEventArgs e)
        //{
        //    if (!_allowClose)
        //    {
        //        e.Cancel = true;
        //        return;
        //    }

        //    base.OnFormClosing(e);
        //}
        /// <summary>
        /// 动态生成SN输入框
        /// </summary>
        private void InitSNInput(int count)
        {
            int marginTop = 60;
            int rowHeight = 45;

            int labelWidth = 60;
            int textWidth = 250;
            int lenWidth = 60;
            int statusWidth = 100;

            int startX = 20;

            for (int i = 0; i < count; i++)
            {
                int y = marginTop + (i * rowHeight);

                // 左侧标题
                UILabel label = new UILabel();
                label.Text = $"SN{i + 1}";
                label.Location = new Point(startX, y + 5);
                label.Size = new Size(labelWidth, 30);

                // 输入框
                UITextBox txt = new UITextBox();
                txt.Name = $"txtSN{i + 1}";
                txt.Location = new Point(startX + labelWidth, y);
                txt.Size = new Size(textWidth, 35);

                // 长度显示
                UILabel lenLabel = new UILabel();
                lenLabel.Text = "0";
                lenLabel.ForeColor = Color.Gray;
                lenLabel.TextAlign = ContentAlignment.MiddleLeft;
                lenLabel.Location = new Point(
                    startX + labelWidth + textWidth + 10,
                    y + 5);

                lenLabel.Size = new Size(lenWidth, 30);

                // 状态显示
                UILabel statusLabel = new UILabel();
                statusLabel.Text = "";
                statusLabel.ForeColor = Color.Gray;
                statusLabel.TextAlign = ContentAlignment.MiddleLeft;

                statusLabel.Location = new Point(
                    startX + labelWidth + textWidth + lenWidth + 20,
                    y + 5);

                statusLabel.Size = new Size(statusWidth, 30);

                // 文本变化
                txt.TextChanged += (s, e) =>
                {
                    string sn = txt.Text.Trim();

                    // 长度
                    lenLabel.Text = sn.Length.ToString();

                    // 默认颜色
                    txt.FillColor = Color.White;

                    statusLabel.Text = "";
                    statusLabel.ForeColor = Color.Gray;

                    // 空值不检查
                    if (string.IsNullOrWhiteSpace(sn))
                        return;

                    // 检查重复
                    int duplicateCount = _textBoxes
                        .Count(t => t.Text.Trim() == sn);

                    if (duplicateCount > 1)
                    {
                        txt.FillColor = Color.MistyRose;

                        statusLabel.Text = "重复";
                        statusLabel.ForeColor = Color.Red;

                        return;
                    }

                    // 检查长度（示例：SN长度必须20）
                    if (sn.Length != _snLength)
                    {
                        txt.FillColor = Color.LightYellow;

                        statusLabel.Text = $"长度错误 SnLen:{sn.Length}  GoalLen: {_snLength}";
                        statusLabel.ForeColor = Color.DarkOrange;

                        return;
                    }
                    //MES
                    if (_mes)
                    {
                        // 执行MES检查逻辑
                        var message = string.Empty;
                        var flag = MES_Service.CheckSerialNumber(sn, ref message);
                        if (!flag)
                        {
                            this.ShowErrorNotifier(message);
                            statusLabel.Text = message;
                            statusLabel.ForeColor = Color.DarkOrange;
                            return;
                        }
                    }
                    // 正常
                    txt.FillColor = Color.Honeydew;

                    statusLabel.Text = "OK";
                    statusLabel.ForeColor = Color.LimeGreen;
                };

                int currentIndex = i;

                // 回车跳转
                txt.KeyDown += (s, e) =>
                {
                    if (e.KeyCode != Keys.Enter)
                        return;

                    e.SuppressKeyPress = true;

                    //  int currentIndex = i;

                    // ❗ 先校验当前输入（关键修改点）
                    var sn = _textBoxes[currentIndex].Text.Trim();

                    // 空
                    if (string.IsNullOrWhiteSpace(sn))
                    {
                        UIMessageTip.ShowError("SN不能为空");
                        txt.Focus();
                        txt.SelectAll();
                        return;
                    }

                    // 长度（建议补上，否则重复/OK状态不可靠）
                    if (sn.Length != _snLength)
                    {
                        UIMessageTip.ShowError("SN长度错误");
                        txt.Focus();
                        txt.SelectAll();
                        return;
                    }

                    // 重复检测
                    int dup = _textBoxes.Count(t => t.Text.Trim() == sn);
                    if (dup > 1)
                    {
                        UIMessageTip.ShowError("SN重复");
                        txt.Focus();
                        txt.SelectAll();
                        return;
                    }

                    // MES校验（如果启用）
                    if (_mes)
                    {
                        var message = string.Empty;
                        if (!MES_Service.CheckSerialNumber(sn, ref message))
                        {
                            UIMessageTip.ShowError(message);
                            txt.Focus();
                            txt.SelectAll();
                            return;
                        }
                    }

                    // ✔ 如果不是最后一个 → 跳转下一个
                    if (currentIndex + 1 < _textBoxes.Count)
                    {
                        _textBoxes[currentIndex + 1].Focus();
                        _textBoxes[currentIndex + 1].SelectAll();
                        return;
                    }

                    // ✔ 最后一个 → 再做全局最终校验（保险）
                    bool allOk = true;

                    for (int ie = 0; ie < _textBoxes.Count; ie++)
                    {
                        var ss = _textBoxes[ie].Text.Trim();

                        if (string.IsNullOrWhiteSpace(ss))
                        {
                            allOk = false;
                            break;
                        }

                        if (_textBoxes.Count(t => t.Text.Trim() == ss) > 1)
                        {
                            allOk = false;
                            break;
                        }

                        if (_statusLabels[ie].Text != "OK")
                        {
                            allOk = false;
                            break;
                        }
                    }

                    if (allOk)
                    {
                        this.DialogResult = DialogResult.OK;
                        this.Close();
                    }
                    else
                    {
                        UIMessageTip.ShowError("存在未通过校验的SN，请检查！");
                    }
                };

                _textBoxes.Add(txt);
                _statusLabels.Add(statusLabel);

                this.Controls.Add(label);
                this.Controls.Add(txt);
                this.Controls.Add(lenLabel);
                this.Controls.Add(statusLabel);
            }

            // 自动调整窗体大小
            int formWidth =
                startX +
                labelWidth +
                textWidth +
                lenWidth +
                statusWidth +
                5;

            int formHeight =
                marginTop +
                (count * rowHeight) +
                10;

            this.Size = new Size(formWidth, formHeight);

            // 禁止缩小
            this.MinimumSize = this.Size;

            // 居中
            this.StartPosition = FormStartPosition.CenterScreen;

            // 默认焦点
            if (_textBoxes.Count > 0)
            {
                this.Shown += (s, e) =>
                {
                    _textBoxes[0].Focus();
                };
            }
        }

        public int GetFilledCount()
        {
            return _textBoxes.Count(t => !string.IsNullOrWhiteSpace(t.Text));
        }

        public List<string> GetAllSN()
        {
            return _textBoxes
                .Select(t => t.Text.Trim())
                .ToList();
        }

        public bool IsAllOk()
        {
            return _statusLabels.All(l => l.Text == "OK");
        }
    }
}