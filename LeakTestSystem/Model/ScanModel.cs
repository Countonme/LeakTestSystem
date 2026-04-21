using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LeakTestSystem.Model
{
    public class ScanModel
    {
        public string serialNumber { get; set; }
        public string msg { get; set; }

        public bool result { get; set; }

        public string error { get; set; }

        public string message { get; set; }

        public bool model { get; set; }
    }
}