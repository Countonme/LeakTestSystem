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
    }
}
