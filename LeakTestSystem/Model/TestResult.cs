using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace LeakTestSystem.Model
{
    public class TestResult
    {
        /// <summary>
        /// 程序号
        /// </summary>
        public string programNo { get; set; }

        /// <summary>
        /// SN
        /// </summary>
        public string serialNumber { get; set; }

        /// <summary>
        /// 测试结果 OK / NG / AL
        /// </summary>
        public string testResult { get; set; }

        /// <summary>
        /// 设备ID
        /// </summary>
        public string equipmentId { get; set; }

        /// <summary>
        /// 设备名称
        /// </summary>
        public string equipmentName { get; set; }

        /// <summary>
        /// 通道名称 CH1 CH2
        /// </summary>
        public string channelsName { get; set; }

        /// <summary>
        /// 气压值
        /// </summary>
        public string PressureValue { get; set; }

        /// <summary>
        /// 泄漏值
        /// </summary>
        public string LeakValue { get; set; }

        /// <summary>
        /// PC 名称
        /// </summary>
        public string pcName { get; set; }

        /// <summary>
        /// IP
        /// </summary>
        public string IpAddress { get; set; }

        /// <summary>
        /// COM名称
        /// </summary>
        public string comName { get; set; }

        /// <summary>
        /// 报警信息
        /// </summary>
        public string alarmMessage { get; set; }

        /// <summary>
        /// 原始数据
        /// </summary>
        public string rawData { get; set; }

        public static bool TryParse(string input, out TestResult result)
        {
            result = null;

            if (string.IsNullOrWhiteSpace(input))
                return false;

            input = input.Trim();

            // ⭐ 统一结构（兼容缺字段）
            string pattern =
                @"<(?<programNo>\d+)>:" +
                @"(?<pressure>-?[\d.]+)?\s*(?<pressureUnit>\w+)?:?" +
                @"\((?<status>OK|TD|RD|NG|AL)\):" +
                @"(?<payload>.*)";

            Match match = Regex.Match(input, pattern);

            if (!match.Success)
                return false;

            string status = match.Groups["status"].Value;
            string payload = match.Groups["payload"].Value;

            result = new TestResult
            {
                rawData = input,
                programNo = match.Groups["programNo"].Value,
                testResult = status
            };

            // ⭐ pressure（可能不存在）
            if (match.Groups["pressure"].Success)
            {
                result.PressureValue =
                    $"{match.Groups["pressure"].Value} {match.Groups["pressureUnit"].Value}".Trim();
            }

            // ⭐ 根据状态处理 payload
            if (status == "OK" || status == "TD" || status == "RD")
            {
                // OK类：payload是流量
                result.LeakValue = payload.Trim();
            }
            else
            {
                // NG / AL：payload是错误信息
                result.alarmMessage = payload.Trim();
            }

            return true;
        }
    }
}