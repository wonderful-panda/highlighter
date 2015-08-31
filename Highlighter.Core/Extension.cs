using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Highlighter.Core
{
    public static class Extension
    {
        public static int ToMsoColor(this System.Drawing.Color color)
        {
            return color.R + color.G * 0x100 + color.B * 0x10000;
        }
    }
}
