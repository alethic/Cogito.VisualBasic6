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

        public static int Main(string[] args)
        {
            try
            {
                int r = 0;
                Parser.Default.ParseArguments<Options>(args).WithParsed(o =>
                {
                    r = Run(o);
                });

                return r;
            }
            catch (Exception e)
            {
                Console.Error.Write(e.ToString());
                return 1;
            }
        }

        static int Run(Options options)
        {
            var e = new Executor()
            {
                Exe = options.Exe.Trim().Trim('"').Trim(),
                Vbp = options.Vbp.Trim().Trim('"').Trim(),
                Out = options.Out.Trim().Trim('"').Trim(),
            };

            return e.Execute(Console.Error);
        }

    }

}
