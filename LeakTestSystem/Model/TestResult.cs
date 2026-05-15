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

            // ⭐ OK 数据（支持负数）
            string okPattern =
                @"<(?<programNo>\d+)>:" +
                @"(?<pressure>-?[\d.]+)\s(?<pressureUnit>\w+):" +
                @"\((?<status>OK|TD|RD|NG)\):" +
                @"(?<flow>-?[\d.]+)\s(?<flowUnit>\w+)";

            // ⭐ AL 数据
            string alarmPattern =
                @"<(?<programNo>\d+)>:" +
                @"\((?<status>AL)\):" +
                @"(?<message>.+)";

            Match match = Regex.Match(input, okPattern);

            if (match.Success)
            {
                result = new TestResult
                {
                    rawData = input,
                    programNo = match.Groups["programNo"].Value,
                    testResult = match.Groups["status"].Value,

                    PressureValue =
                        $"{match.Groups["pressure"].Value} {match.Groups["pressureUnit"].Value}",

                    LeakValue =
                        $"{match.Groups["flow"].Value} {match.Groups["flowUnit"].Value}"
                };

                return true;
            }

            match = Regex.Match(input, alarmPattern);

            if (match.Success)
            {
                result = new TestResult
                {
                    rawData = input,
                    programNo = match.Groups["programNo"].Value,
                    testResult = match.Groups["status"].Value,
                    alarmMessage = match.Groups["message"].Value
                };

                return true;
            }

            return false;
        }
    }
}