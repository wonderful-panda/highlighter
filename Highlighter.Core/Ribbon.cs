using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using Office = Microsoft.Office.Core;
using ColorCode;

namespace Highlighter.Core
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        readonly IHighlighterHost _host;
        readonly ColorCode.ILanguage []_langs;
        ColorCode.ILanguage _defaultLang;

        private Office.IRibbonUI ribbon;

        public Ribbon(IHighlighterHost host)
        {
            _host = host;
            _langs = ColorCode.Languages.All.OrderBy(lang => lang.Name).ToArray();
            _defaultLang = _langs[0];
        }

        public string GetCustomUI(string ribbonID)
        {
            var xml = new RibbonXml(_langs).TransformText();
            return xml;
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #region Callbacks

        /// <summary>
        /// Get label of SplitButton
        /// </summary>
        /// <param name="control"></param>
        /// <returns></returns>
        public string GetButtonLabel(Office.IRibbonControl control)
        {
            return "As " + _defaultLang.Name;
        }

        /// <summary>
        /// Called when SplitButton or it's submenus are clicked.
        /// </summary>
        /// <param name="control"></param>
        public void OnHighlightAction(Office.IRibbonControl control)
        {
            if (!string.IsNullOrEmpty(control.Tag))
                _defaultLang = ColorCode.Languages.FindById(control.Tag);
            _host.Highlight(new Parser(_defaultLang));
            this.ribbon.Invalidate();
        }

        #endregion
    }
}
