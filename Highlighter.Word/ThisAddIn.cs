using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using WordInterop = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Highlighter.Core;

namespace Highlighter.Word
{
    public partial class ThisAddIn : Highlighter.Core.IHighlighterHost
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

        public void Highlight(Core.Parser parser)
        {
            this.Application.ScreenUpdating = false;
            try
            {
                var sel = this.Application.Selection;
                if (!string.IsNullOrEmpty(sel.Text))
                {
                    // Apply highlight to selected text
                    HighlightRange(sel.Range, parser);
                }
                else
                {
                    // Apply highlight to selected shapes
                    foreach (WordInterop.Shape s in sel.ShapeRange)
                    {
                        if (s.TextFrame.HasText != 0)
                            HighlightRange(s.TextFrame.TextRange, parser);
                    }
                }
            }
            finally
            {
                this.Application.ScreenUpdating = true;
            }
        }

        void HighlightRange(WordInterop.Range range, Core.Parser parser)
        {
            // reset text color.
            range.Font.ColorIndex = WordInterop.WdColorIndex.wdAuto;

            var chars = range.Characters;
            foreach (var f in parser.Parse(range.Text))
            {
                // NOTE: indices of Characters are 1-based
                var r = chars[1 + f.Start];
                r.SetRange(r.Start, r.Start + f.Length);
                r.Font.Fill.ForeColor.RGB = f.ForegroundColor.ToMsoColor();
            }
        }
    }
}
