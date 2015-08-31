using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Highlighter.Core
{
    public interface IHighlighterHost
    {
        void Highlight(Parser parser);
    }
}
