using System.Collections.Generic;
using System.Linq;

using CommandLine;

namespace Cogito.VisualBasic6.Make
{

    public static class Program
    {

        class Options
        {

            [Option('e', "vb6", Required = true, HelpText = "Path to the original VB6.exe.")]
            public string VB6 { get; set; }

            [Option('m', "make", Required = true, HelpText = "Path to the .vbp file to be built.")]
            public string Make { get; set; }

            [Option('o', "outdir", Default = ".", HelpText = "Path to output directory.")]
            public string OutDir { get; set; }

            [Option('d', "define", Separator = ':', HelpText = "Colon separated list of name=value pairs to be defined.")]
            public IEnumerable<string> Define { get; set; }

            [Value(0, HelpText = "Path to the output file to be built.")]
            public string Target { get; set; }

        }

        public static void Main(string[] args)
        {
            Parser.Default.ParseArguments<Options>(args).WithParsed(Run);
        }

        static void Run(Options options)
        {
            var exec = new Executor()
            {
                Exe = options.VB6,
                Vbp = options.Make,
                Out = options.Target,
                Dir = options.OutDir,
                Def = options.Define?.Select(i => i.Split(new[] { '=' }, 2)).ToDictionary(i => i[0], i => i[1]),
            };

            exec.Execute();
        }

    }

}
