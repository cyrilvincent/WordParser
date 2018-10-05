using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace WordParserLibrary
{
    public class PreCacheParser<SCOPE> where SCOPE : IScope, new()
    {
        public DocX Doc { get; set; }
        public string TemplateFile { get; set; } = "Exports/Word/Template.docx";
        public string TempFile { get; set; } = "Exports/Word/Temp/Template.Precompile.docx";
        public List<Expression<SCOPE>> Expressions { get; set; }
        public SCOPE Scope { get; }
        public int NbIteration { get; set; } = 5;

        private RegexOptions regexOptions = RegexOptions.CultureInvariant;

        public PreCacheParser(SCOPE scope)
        {
            Scope = scope;
        }

        public static void PreCache()
        {
            SCOPE scope = new SCOPE();
            PreCacheParser<SCOPE> precache = new PreCacheParser<SCOPE>(scope);
            precache.Parse();
            precache.Compiles();
        }

        public void Parse()
        {
            File.Copy(TemplateFile, TempFile, true);
            Doc = DocX.Load(TempFile);
            string re = "{{.+?}}";
            List<string> es = Doc.FindUniqueByPattern(re, regexOptions);
            Expressions = es.Select(s => Expression<SCOPE>.Factory(s, new SCOPE())).ToList();
            Regex regex = new Regex(@"{{(.+)\[i\]");
            foreach (Expression<SCOPE> e in Expressions.ToList())
            {
                Match m = regex.Match(e.Raw);
                if (m.Success)
                {
                    string code = m.Groups[1].Value;
                    code += ".Count";
                    Expressions.Add(new Expression<SCOPE>(Scope) { Code = code });
                    for (int i = 0; i <= NbIteration - 1; i++)
                    {
                        code = e.Code.Replace("[i]", "[" + i + "]");
                        Expressions.Add(new Expression<SCOPE>(Scope) { Code = code });
                    }
                }
            }
        }

        public void Compiles()
        {
            foreach (Expression<SCOPE> e in Expressions)
            {
                e.Scope = Scope; // Certainement inutile
                Compile(e);
            }
        }

        public void Compile(Expression<SCOPE> e)
        {
            try
            {
                e.Compile();
            }
            catch (AggregateException ex)
            {
                if (ex.InnerException != null)
                    e.Lambda = s => "{{CACHE COMPILE ERROR " + e.Code + " " + ex.InnerException.Message + "}}";
                else
                    e.Lambda = s => "{{CACHE COMPILE AGGREGRATED ERROR " + e.Code + " " + ex.Message + "}}";
            }
            catch (Exception ex)
            {
                e.Lambda = s => "{{CACHE COMPILE UNKNOWN ERROR " + e.Code + " " + ex.Message + "}}";
            }
        }
    }
}
