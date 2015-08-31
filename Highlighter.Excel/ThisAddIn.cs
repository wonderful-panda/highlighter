using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Highlighter.Core;

namespace Highlighter.Excel
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
            try
            {
                this.Application.ScreenUpdating = false;
                var sel = this.Application.Selection;
                if (sel is ExcelInterop.Range)
                {
                    // Highlight cells
                    HighlightCells(sel, parser);
                }
                else
                {
                    // Highlight shapes
                    foreach (ExcelInterop.Shape shape in sel.ShapeRange)
                    {
                        if (shape.TextFrame2.HasText == Office.MsoTriState.msoTrue)
                            HighlightTextRange(shape.TextFrame2.TextRange, parser);
                    }
                }
            }
            finally
            {
                this.Application.ScreenUpdating = true;
            }
        }

        void HighlightCells(ExcelInterop.Range range, Core.Parser parser)
        {
            if (range.Cells.Count > 1)
            {
                foreach (var cell in range.Cells.Cast<ExcelInterop.Range>())
                {
                    HighlightCells(cell, parser);
                }
            }
            else
            {
                var text = (string)range.Text;
                if (string.IsNullOrEmpty(text))
                    return;
                // reset text color.
                range.Font.ColorIndex = ExcelInterop.XlColorIndex.xlColorIndexAutomatic;

                foreach (var f in parser.Parse(text))
                {
                    // NOTE: indices of Characters are 1-based
                    range.get_Characters(f.Start + 1, f.Length).Font.Color = f.ForegroundColor;
                }
            }
        }

        void HighlightTextRange(Office.TextRange2 textRange, Core.Parser parser)
        {
            var text = textRange.Text;
            if (string.IsNullOrEmpty(text))
                return;
            // reset text color.
            textRange.Font.Fill.ForeColor.RGB = System.Drawing.Color.Black.ToMsoColor();

            foreach (var f in parser.Parse(text))
            {
                // NOTE: indices of Characters are 1-based
                textRange.get_Characters(1 + f.Start, f.Length).Font.Fill.ForeColor.RGB = f.ForegroundColor.ToMsoColor();
            }
        }
    }
}
