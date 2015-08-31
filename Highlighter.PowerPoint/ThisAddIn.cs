using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPointInterop = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Highlighter.Core;

namespace Highlighter.PowerPoint
{
    public partial class ThisAddIn : Core.IHighlighterHost
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Highlighter.Core.Ribbon(this);
        }

        public void Highlight(Parser parser)
        {
            var sel = this.Application.ActiveWindow.Selection;
            if (sel.ShapeRange.Count == 1)
            {
                    HighlightTextRange(sel.TextRange2, parser);
            }
            else
            {
                foreach (PowerPointInterop.Shape shape in sel.ShapeRange)
                {
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                        HighlightTextRange(shape.TextFrame2.TextRange, parser);
                }

            }
        }

        void HighlightTextRange(Office.TextRange2 textRange, Core.Parser parser)
        {
            textRange.Font.Fill.ForeColor.RGB = System.Drawing.Color.Black.ToMsoColor();
            foreach (var f in parser.Parse(textRange.Text))
            {
                // NOTE: indices of Characters are 1-based
                textRange.get_Characters(1 + f.Start, f.Length).Font.Fill.ForeColor.RGB = f.ForegroundColor.ToMsoColor();
            }
        }
    }
}
