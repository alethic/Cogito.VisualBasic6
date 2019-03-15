using System;
using System.Collections.Generic;
using System.IO;
using Cogito.VisualBasic6.VB6C.Project;
using CommandLine;

namespace Cogito.VisualBasic6.VB6C
{

    public static class Program
    {

        class Options
        {

            [Option('e', "vb6", Required = true, HelpText = "Path to the original VB6.exe.")]
            public string VB6 { get; set; }

            [Option('s', "source", HelpText = "Path to the source .vbp file to be imported.")]
            public string Source { get; set; }

            [Option('t', "target", Default = ".", HelpText = "Path to output file.")]
            public string Target { get; set; }

            [Option('o', "output", Default = ".", HelpText = "Path to output directory.")]
            public string Output { get; set; }

            [Option('r', "references", Separator = ';', HelpText = "Paths to COM object references.")]
            public IEnumerable<string> References { get; set; }

            [Option('p', "properties", Separator = ':', HelpText = "Colon separated list of additional name=value pairs to be defined.")]
            public IEnumerable<string> Properties { get; set; }

            [Value(0, HelpText = "Input files to be bult.")]
            public string Input { get; set; }

        }

        public static void Main(string[] args)
        {
            new Compiler().Compile(
                new FileInfo(@"c:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE"),
                VB6Project.Load(@"C:\dev\Cogito.VisualBasic6\sample\Cogito.VisualBasic6.Sample\obj\Debug\net47\VB6Sample.vbp"),
                new DirectoryInfo(@"C:\dev\Cogito.VisualBasic6\sample\Cogito.VisualBasic6.Sample\obj\Debug\net47"),
                Console.Out);
        }

    }

}
