using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using WordParserLibrary;

namespace WordParserConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Devis Machine Console");
            Console.WriteLine("=====================");

            Console.WriteLine("Caching...");
            Thread thread = new Thread(PreCacheParser<MockScope>.PreCache);
            thread.Priority = ThreadPriority.AboveNormal;
            thread.Start();

            Console.WriteLine("Scope creation...");
            MockScope scope = new MockScope { Id = 1, Text = "Hello World", FileName = "WordParser" };
            for(int i = 0; i< 10; i++)
            {
                scope.List.Add(new ABC { Label = "Label"+i.ToString(), A = "A" + i.ToString(), B = "B" + i.ToString(), C = "C" + i.ToString() });
            }

            while (true)
            {
                WordParser<MockScope> parser = new WordParser<MockScope>(scope);
                Console.WriteLine("Parsing...");
                parser.Parse();
                int nbCache = 0;
                if(Expression<MockScope>.Cache != null)
                {
                    nbCache = Expression<MockScope>.Cache.Count;
                }
                Console.WriteLine("Mapping " + parser.Expressions.Count + " expression(s) with "+nbCache+" item(s) in cache");
                parser.Map();
                Console.WriteLine("Saving...");
                parser.Save();
                string fileName = Directory.GetCurrentDirectory() + "\\" + parser.TempFile;
                Console.WriteLine("Opening " + fileName);
                Process process = new Process();
                process.StartInfo.FileName = fileName;
                process.Start();
                if(parser.NbError > 0)
                {
                    Console.WriteLine(parser.NbError + " ERROR(S)");
                }
                Console.WriteLine("Press any key...");
                Console.ReadKey();
                Console.WriteLine();
            }
        }
    }
}
