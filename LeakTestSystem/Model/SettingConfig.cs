using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LeakTestSystem.Model
{
    public class SettingConfig
    {
        /// <summary>
        /// 对应的Com 名称
        /// </summary>
        public string masterComName { get; set; }

        /// <summary>
        /// 对应的Com 名称
        /// </summary>
        public string channel1ComName { get; set; }

        /// <summary>
        /// 1 Enable, 0 Disable
        /// </summary>
        public bool channel1Status { get; set; }

        /// <summary>
        ///  channel 2 Index ##########################################
        /// <summary>
        /// 对应的Com 名称
        /// </summary>
        public string channel2ComName { get; set; }

        /// <summary>
        /// 1 Enable, 0 Disable
        /// </summary> 
        public bool channel2Status { get; set; }

        /// <summary>
        ///  channel 3 Index ##########################################
        /// <summary>
        /// 对应的Com 名称
        /// </summary>
        public string channel3ComName { get; set; }

        /// <summary>
        /// 1 Enable, 0 Disable
        /// </summary> 
        public bool channel3Status { get; set; }



        /// <summary>
        ///  channel 4 Index ##########################################
        /// <summary>
        /// 对应的Com 名称
        /// </summary>
        public string channel4ComName { get; set; }

        /// <summary>
        /// 1 Enable, 0 Disable
        /// </summary> 
        public bool channel4Status { get; set; }


        /// <summary>
        ///  channel 5 Index ##########################################
        /// <summary>
        /// 对应的Com 名称
        /// </summary>
        public string channel5ComName { get; set; }

        /// <summary>
        /// 1 Enable, 0 Disable
        /// </summary> 
        public bool channel5Status { get; set; }

        /// <summary>
        ///  channel 6 Index #########################################
        /// <summary>
        /// 对应的Com 名称
        /// </summary>
        public string channel6ComName { get; set; }

        /// <summary>
        /// 1 Enable, 0 Disable
        /// </summary> 
        public bool channel6Status { get; set; }
        /// <summary>
        /// 读取超时
        /// </summary>
        public int readTimeout { get; set; } = 20;
        /// <summary>
        /// MES NG 锁码
        /// </summary>
        public bool mesNgLock { get; set; }
        /// <summary>
        /// SN length
        /// </summary>
        public int snLength { get; set; } = 6;  
        public int GetEnableChannelCount(SettingConfig config)
        {
            int count = 0;

            if (config.channel1Status) count++;
            if (config.channel2Status) count++;
            if (config.channel3Status) count++;
            if (config.channel4Status) count++;
            if (config.channel5Status) count++;
            if (config.channel6Status) count++;

            return count;
        }


        public int GetChannelIndexByComName(SettingConfig config, string comName)
        {
            var enableChannels = new List<string>();

            if (config.channel1ComName==comName) return 0;
            if (config.channel2ComName==comName) return 1;
            if (config.channel3ComName==comName) return 2;
            if (config.channel4ComName==comName) return 3;
            if (config.channel5ComName==comName) return 4;
            if (config.channel6ComName==comName) return 5;

            return -1;
        }

        public string GetEnableChannelName(SettingConfig config, int index)
        {
            var enableChannels = new List<string>();

            if (config.channel1Status) enableChannels.Add("channel1");
            if (config.channel2Status) enableChannels.Add("channel2");
            if (config.channel3Status) enableChannels.Add("channel3");
            if (config.channel4Status) enableChannels.Add("channel4");
            if (config.channel5Status) enableChannels.Add("channel5");
            if (config.channel6Status) enableChannels.Add("channel6");

            if (index < 0 || index >= enableChannels.Count)
                return null;

            return enableChannels[index];
        }

        public int GetEnableChannelIndex(SettingConfig config, int index)
        {
            var enableChannels = new List<int>();

            if (config.channel1Status) enableChannels.Add(1);
            if (config.channel2Status) enableChannels.Add(2);
            if (config.channel3Status) enableChannels.Add(3);
            if (config.channel4Status) enableChannels.Add(4);
            if (config.channel5Status) enableChannels.Add(5);
            if (config.channel6Status) enableChannels.Add(6);

            if (index < 0 || index >= enableChannels.Count)
                return -1;

            return enableChannels[index];
        }
        public int GetEnableChannelIndexByCom(string comName)
        {
            var channels = new List<(string com, int index)>
            {
                (channel1ComName, 0),
                (channel2ComName, 1),
                (channel3ComName, 2),
                (channel4ComName, 3),
                (channel5ComName, 4),
                (channel6ComName, 5),
            };

            int enableIndex = 0;

            foreach (var ch in channels)
            {
                if (!string.IsNullOrEmpty(ch.com))
                {
                    if (IsEnabled(ch.index))
                    {
                        enableIndex++;

                        if (ch.com == comName)
                            return enableIndex - 1; // 从0开始
                    }
                }
            }

            return -1;
        }

        private bool IsEnabled(int index)
        {
            switch (index)
            {
                case 0: return channel1Status;
                case 1: return channel2Status;
                case 2: return channel3Status;
                case 3: return channel4Status;
                case 4: return channel5Status;
                case 5: return channel6Status;
                default: return false;
            }
            ;
        }
    }
}
