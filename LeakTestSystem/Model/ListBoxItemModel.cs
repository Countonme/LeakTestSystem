using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace LeakTestSystem.Model
{
    public class ListBoxItemModel
    {
        public string Text { get; set; }

        public Color Color { get; set; }

        public override string ToString()
        {
            return Text;
        }
    }
}