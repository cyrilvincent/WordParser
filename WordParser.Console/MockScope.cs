using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordParserLibrary;

namespace WordParserConsole
{
    public class MockScope : IScope
    {
        public int Id { get; set; }
        public string FileName { get; set; }
        public string Text { get; set; }
        public List<ABC> List { get; set; } = new List<ABC>();
    }

    public class ABC
    {
        public string A { get; set; }
        public string B { get; set; }
        public string C { get; set; }
        public string Label { get; set; }
    }
}
