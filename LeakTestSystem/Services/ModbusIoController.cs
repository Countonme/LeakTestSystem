using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LeakTestSystem.Services
{
    public class ModbusIoController
    {
        private SerialPort _port;

        public ModbusIoController(SerialPort port)
        {
            _port = port;
            if (_port == null)
                throw new Exception("SerialPort 不能为空");
        }

        /// <summary>
        /// 控制继电器（闭合/断开）
        /// </summary>
        /// <param name="slaveId">从站地址</param>
        /// <param name="coil">线圈地址（0开始）</param>
        /// <param name="close">true=闭合 false=断开</param>
        public void SetRelay(int slaveId, int coil, bool close)
        {
            byte[] frame = Build05Frame(slaveId, coil, close);

            if (_port != null && _port.IsOpen)
            {
                _port.Write(frame, 0, frame.Length);
            }
            else
            {
                throw new Exception("串口未打开");
            }
        }

        /// <summary>
        /// 闭合
        /// </summary>
        public void Close(int slaveId, int coil)
        {
            SetRelay(slaveId, coil, true);
        }

        /// <summary>
        /// 断开
        /// </summary>
        public void Open(int slaveId, int coil)
        {
            SetRelay(slaveId, coil, false);
        }

        /// <summary>
        /// 构建 05 功能码报文
        /// </summary>
        private byte[] Build05Frame(int slaveId, int coil, bool close)
        {
            List<byte> data = new List<byte>();

            data.Add((byte)slaveId);
            data.Add(0x05);

            // 线圈地址
            data.Add((byte)(coil >> 8));
            data.Add((byte)(coil & 0xFF));

            // 状态
            if (close)
            {
                data.Add(0xFF);
                data.Add(0x00);
            }
            else
            {
                data.Add(0x00);
                data.Add(0x00);
            }

            ushort crc = CalcCRC(data.ToArray());
            data.Add((byte)(crc & 0xFF));
            data.Add((byte)(crc >> 8));

            return data.ToArray();
        }

        /// <summary>
        /// Modbus CRC16
        /// </summary>
        private ushort CalcCRC(byte[] data)
        {
            ushort crc = 0xFFFF;

            for (int i = 0; i < data.Length; i++)
            {
                crc ^= data[i];

                for (int j = 0; j < 8; j++)
                {
                    if ((crc & 1) != 0)
                    {
                        crc >>= 1;
                        crc ^= 0xA001;
                    }
                    else
                    {
                        crc >>= 1;
                    }
                }
            }

            return crc;
        }
    }
}
