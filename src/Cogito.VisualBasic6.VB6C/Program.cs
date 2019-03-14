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

            [Option('i', "project", HelpText = "Path to the source .vbp file to be imported.")]
            public string Project { get; set; }

            [Option('o', "out", Default = ".", HelpText = "Path to output file.")]
            public string Out { get; set; }

            [Option('d', "dir", Default = ".", HelpText = "Path to output directory.")]
            public string Dir { get; set; }

            [Option('p', "properties", Separator = ':', HelpText = "Colon separated list of additional name=value pairs to be defined.")]
            public IEnumerable<string> Properties { get; set; }

            [Value(0, HelpText = "Input files to be bult.")]
            public string Input { get; set; }

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
                Vbp = options.Project,
                Out = options.Input,
                Dir = options.Dir,
                Def = options.Properties?.Select(i => i.Split(new[] { '=' }, 2)).ToDictionary(i => i[0], i => i[1]),
            };

            exec.Execute();
        }

    }

}
