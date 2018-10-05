using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.Scripting;

namespace WordParserLibrary
{
    public class Expression<SCOPE> where SCOPE : IScope
    {
        public string Code { get; set; }
        public SCOPE Scope { get; set; }
        public static Dictionary<string, Func<SCOPE, object>> Cache { get; set; }

        public string Raw
        {
            get
            {
                return "{{" + Code + "}}";
            }
            set
            {
                Code = value.Substring(2, value.Length - 4);
            }
        }

        public string LambdaString
        {
            get
            {
                return "scope => " + Code;
            }
        }

        public object Value { get; set; } = "null";

        public Expression(SCOPE scope)
        {
            if (Cache == null)
                Cache = new Dictionary<string, Func<SCOPE, object>>();
            this.Scope = scope;
        }

        public static Expression<SCOPE> Factory(string raw, SCOPE scope)
        {
            Expression<SCOPE> e = new Expression<SCOPE>(scope);
            e.Raw = raw;
            return e;
        }

        public override string ToString()
        {
            return Code + " => " + Value;
        }

        private ScriptOptions options = ScriptOptions.Default.AddReferences(typeof(SCOPE).Assembly);

        public Func<SCOPE, object> Lambda { get; set; }

        public async Task<T> CompileAsync<T>()
        {
            T result = default(T);
            result = await CSharpScript.EvaluateAsync<T>(LambdaString, options);
            return result;
        }

        public async Task<Func<SCOPE, object>> CompileAsync()
        {
            Func<SCOPE, object> result = await CompileAsync<Func<SCOPE, object>>();
            return result;
        }

        public void Compile()
        {
            if (Cache.ContainsKey(LambdaString))
                Lambda = Cache[LambdaString];
            else
            {
                Task<Func<SCOPE, object>> result = CompileAsync();
                result.Wait();
                Lambda = result.Result;
                try
                {
                    Cache[LambdaString] = Lambda;
                }
                catch (Exception)
                {
                }// For Multithreading
            }
        }
    }
}
