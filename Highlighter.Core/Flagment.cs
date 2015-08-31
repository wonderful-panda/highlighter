using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace Highlighter.Core
{
    public class Flagment
    {
        public string Name { get; private set; }
        public string Value { get; private set; }
        public Color ForegroundColor { get; private set; }
        public int Start { get; private set; }
        public int Length { get; private set; }

        public Flagment(string name, string value, Color foregroundColor, int start, int length)
        {
            this.Name = name;
            this.Value = value;
            this.ForegroundColor = foregroundColor;
            this.Start = start;
            this.Length = length;
        }

        internal static Flagment FromScope(int baseIndex, string token, ColorCode.Parsing.Scope scope, ColorCode.IStyleSheet stylesheet)
        {
            return new Flagment(scope.Name, token.Substring(scope.Index, scope.Length), stylesheet.Styles[scope.Name].Foreground, 
                                baseIndex + scope.Index, scope.Length);
        }
    }
}
