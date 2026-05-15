using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LeakTestSystem.Model
{
    public class TestStep
    {
        public int ch { get; set; }
        public string serialNumber { get; set; }

        public string data { get; set; }

        public string timedout { get; set; }
    }
}