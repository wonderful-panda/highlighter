using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Highlighter.Core
{
    /// <summary>
    /// Wrapper of ColorCode.Parsing.LanguageParser
    /// </summary>
    public class Parser
    {
        readonly ColorCode.ILanguage _lang;
        readonly ColorCode.Parsing.LanguageParser _internalParser;

        internal Parser(ColorCode.ILanguage lang)
        {
            _lang = lang;
            var repo = new ColorCode.Common.LanguageRepository(new Dictionary<string, ColorCode.ILanguage>());
            var compiler = new ColorCode.Compilation.LanguageCompiler(new Dictionary<string, ColorCode.Compilation.CompiledLanguage>());
            _internalParser = new ColorCode.Parsing.LanguageParser(compiler, repo);
        }

        public Flagment[] Parse(string code)
        {
            var flagments = new List<Flagment>();
            var baseIndex = 0;
            _internalParser.Parse(code, _lang, (token, scopes) => {
                flagments.AddRange(from s in scopes select Flagment.FromScope(baseIndex, token, s, ColorCode.StyleSheets.Default));
                baseIndex += token.Length;
            });
            return flagments.ToArray();
        }
    }
}
