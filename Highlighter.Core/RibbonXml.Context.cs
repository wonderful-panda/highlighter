using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Highlighter.Core
{
    partial class RibbonXml
    {
        public ColorCode.ILanguage []Languages { get; private set; } 
        public RibbonXml(ColorCode.ILanguage []langs)
        {
            this.Languages = langs;
        }
    }
}
