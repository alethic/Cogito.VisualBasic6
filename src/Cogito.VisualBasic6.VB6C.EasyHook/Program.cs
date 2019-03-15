using System;
using System.IO;

using CommandLine;

namespace Cogito.VisualBasic6.VB6C.EasyHook
{

    public static class Program
    {

        class Options
        {

            [Option('e')]
            public string Exe { get; set; }

            [Option('v')]
            public string Vbp { get; set; }

            [Option('o')]
            public string Out { get; set; }

        }

        public static void Main(string[] args)
        {
            Parser.Default.ParseArguments<Options>(args).WithParsed(Run);
        }

        static void Run(Options options)
        {
            var e = new Executor()
            {
                Exe = options.Exe.Trim().Trim('"').Trim(),
                Vbp = options.Vbp.Trim().Trim('"').Trim(),
                Out = options.Out.Trim().Trim('"').Trim(),
            };

            using (var stderr = Console.OpenStandardError())
                e.Execute(new StreamWriter(stderr));
        }

    }

}
