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
    public class WordParser<SCOPE> where SCOPE : IScope, new()
    {
        public string TemplateFile { get; set; } = "Exports/Word/Template.docx";
        public string TempDirectory { get; set; } = "Exports/Word/Temp";
        public string OutputDirectory { get; set; } = "Exports/Word/Outputs";
        public string OutputFile { get; set; }
        public DocX Doc { get; set; }
        public string TempFile { get; set; }
        public List<Expression<SCOPE>> Expressions { get; set; }
        public List<Image> OptionImage { get; set; } = new List<Image>();
        public int NbError { get; set; } = 0;

        public SCOPE Scope { get; set; }

        private RegexOptions reOptions = RegexOptions.CultureInvariant;

        public WordParser(SCOPE scope)
        {
            string dt = DateTime.Now.ToString("yyMMdd-HHmmss");
            OutputFile = OutputDirectory + "/" + scope.FileName + "-" + scope.Id + "-" + dt + ".docx";
            TempFile = TempDirectory + "/" + scope.FileName + "-" + dt + ".tmp.docx";
            Scope = scope;
        }

        public void CreateTemp()
        {
            if (!Directory.Exists(TempDirectory))
                Directory.CreateDirectory(TempDirectory);
            File.Copy(TemplateFile, TempFile, true);
            if (!Directory.Exists(OutputDirectory))
                Directory.CreateDirectory(OutputDirectory);
        }

        public void CreateOutput()
        {
            if (!Directory.Exists(OutputDirectory))
                Directory.CreateDirectory(OutputDirectory);
            File.Copy(TempFile, OutputFile, true);
            File.Delete(TempFile);
        }

        public void Parse()
        {
            NbError = 0;
            CreateTemp();
            Doc = DocX.Load(TempFile);
            LoadImages();
            MapTables();
            MapDoubleParagraphs();
            string re = "{{.+?}}";
            List<string> es = Doc.FindUniqueByPattern(re, reOptions);
            Expressions = es.Select(s => Expression<SCOPE>.Factory(s, new SCOPE())).ToList();
        }

        public void LoadImages()
        {
        }

        public void Map()
        {
            foreach (Expression<SCOPE> e in Expressions)
            {
                e.Scope = Scope; // Certainement inutile
                Compile(e);
                if (e.Value != null)
                    Doc.ReplaceText(e.Raw, e.Value.ToString());
            }
            MapImages();
            RemoveNulls();
            ParagraphsSplitter();
        }

        public void MapTables()
        {
            Regex re = new Regex(@"{{(.+)\[i\]");
            bool success = false;
            var code = "";
            Match m = null;
            int templateRowIndex = 0;
            foreach (Table table in Doc.Tables)
            {
                try
                {
                    // Par défaut je vais chercher le template à cloner sur la dernière ligne, pose problème au chapitre 4 : solution aller chercher l'info ligne 1 et si absente ligne 2
                    templateRowIndex = table.Rows.Count - 1;
                    code = table.Rows[templateRowIndex].Cells[0].Paragraphs[0].Text;
                    m = re.Match(code);
                    success = m.Success;
                }
                catch (Exception)
                {
                }
                if (success)
                {
                    code = m.Groups[1].Value;
                    code += ".Count";
                    Expression<SCOPE> e = new Expression<SCOPE>(Scope) { Code = code };
                    int length = 0;
                    try
                    {
                        e.Compile();
                        length = (int)e.Lambda(Scope);
                    }
                    catch (Exception ex)
                    {
                        table.Rows[1].Cells[0].InsertParagraph("TABLE ERROR " + code + " " + ex.Message);
                        NbError += 1;
                    }
                    for (int i = 0; i <= length - 2; i++)
                        table.InsertRow(table.Rows[templateRowIndex]);
                    for (int i = 0; i <= length - 1; i++)
                        table.Rows[i + templateRowIndex].ReplaceText("[i]", "[" + i + "]");
                }
            }
        }

        public void MapImages()
        {
            Regex re = new Regex(@"{{(OptionImage)\((\d)\)}}");
            foreach (Table t in Doc.Tables)
            {
                foreach (Row row in t.Rows)
                {
                    foreach (Cell cell in row.Cells)
                    {
                        try
                        {
                            string s = cell.Paragraphs[0].Text;
                            Match match = re.Match(s);
                            if (match.Success)
                                MapImage(cell, int.Parse(match.Groups[2].Value));
                        }
                        catch (Exception)
                        {
                        }
                    }
                }
            }
        }

        private void MapImage(Cell cell, int index)
        {
            Picture picture = OptionImage[index].CreatePicture(14, 14);
            cell.Paragraphs[0].AppendPicture(picture);
            cell.Paragraphs[0].RemoveText(0);
        }

        public void MapDoubleParagraphs()
        {
            Regex re = new Regex(@"{{(.+)\[i\]");
            bool success = false;
            var code = "";
            Match m;
            var count = Doc.Paragraphs.Count;
            int i = 0;
            while (i < count)
            {
                Paragraph p = Doc.Paragraphs[i];
                if (p.Text.Contains("{{scope.Mark(\"DoubleParagraph\")}}"))
                {
                    Doc.RemoveParagraph(p);
                    count -= 1;
                    p = Doc.Paragraphs[i];
                    code = p.Text;
                    m = re.Match(code);
                    success = m.Success;
                    if (success)
                    {
                        Paragraph subp = Doc.Paragraphs[i + 1];
                        code = m.Groups[1].Value;
                        code += ".Count";
                        Expression<SCOPE> e = new Expression<SCOPE>(Scope) { Code = code };
                        int length = 0;
                        try
                        {
                            e.Compile();
                            length = (int)e.Lambda(Scope);
                        }
                        catch (Exception ex)
                        {
                            p.InsertParagraphAfterSelf("COUNT ERROR " + code + " " + ex.Message);
                            NbError += 1;
                        }
                        for (int j = length - 2; j >= 0; j += -1)
                        {
                            Paragraph newp = subp.InsertParagraphAfterSelf(p);
                            count += 1;
                            newp.ReplaceText("[i]", "[" + (j + 1) + "]");
                            Paragraph newsubp = newp.InsertParagraphAfterSelf(subp);
                            count += 1;
                            newsubp.ReplaceText("[i]", "[" + (j + 1) + "]");
                        }
                        p.ReplaceText("[i]", "[0]");
                        subp.ReplaceText("[i]", "[0]");
                    }
                }
                i += 1;
            }
        }

        public Table FindTable(string name)
        {
            int index = Doc.FindAll(name)[0];
            return Doc.Tables.Where(t => t.Index < index).OrderByDescending(t => t.Index).First();
        }

        public void RemoveNulls()
        {
            foreach (Paragraph p in Doc.Paragraphs.ToList())
            {
                if (p.Text.Trim().StartsWith("{{null}}"))
                    Doc.RemoveParagraph(p);
                else if (p.Text.Contains("{{null}}"))
                    p.ReplaceText("{{null}}", "");
            }
        }

        public void ParagraphsSplitter()
        {
            List<Paragraph> l = new List<Paragraph>();
            l = Doc.Paragraphs.Where(p => p.Text.Contains("\\n")).ToList();
            foreach (Paragraph p in l)
            {
                string[] texts = p.Text.Split('\n');
                for (int i = texts.Count() - 1; i >= 0; i += -1)
                {
                    var text = texts[i];
                    if (text == null)
                        text = "";
                    if (i > 0)
                    {
                        if (text == "n")
                            text = "";
                        else
                            text = text.Substring(1);
                    }
                    p.InsertParagraphAfterSelf(text);
                }
                Doc.RemoveParagraph(p);
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
                    e.Lambda = s => "{{COMPILE ERROR " + e.Code + " " + ex.InnerException.Message + "}}";
                else
                    e.Lambda = s => "{{COMPILE AGGREGRATED ERROR " + e.Code + " " + ex.Message + "}}";
                NbError += 1;
            }
            catch (Exception ex)
            {
                e.Lambda = s => "{{COMPILE UNKNOWN ERROR " + e.Code + " " + ex.Message + "}}";
                NbError += 1;
            }
            try
            {
                e.Value = e.Lambda(Scope);
            }
            catch (NullReferenceException)
            {
                e.Value = "{{null}}";
            }
            catch (Exception ex)
            {
                e.Value = "{{VALUE ERROR " + e.Code + " " + ex.Message + "}}";
                NbError += 1;
            }
            if (e.Value == null)
                e.Value = "{{null}}";
        }

        public void Save()
        {
            Doc.Save();
            Doc = DocX.Load(TempFile);
            Doc.SaveAs(OutputFile);
        }
    }
}
