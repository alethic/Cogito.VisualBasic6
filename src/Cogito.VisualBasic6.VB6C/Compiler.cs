using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

using Cogito.VisualBasic6.VB6C.EasyHook;
using Cogito.VisualBasic6.VB6C.Project;

namespace Cogito.VisualBasic6.VB6C
{

    /// <summary>
    /// Compiles a <see cref="VB6Project"/>.
    /// </summary>
    public class Compiler
    {

        /// <summary>
        /// Initializes a new instnace.
        /// </summary>
        /// <param name="vb6exe"></param>
        public Compiler(string vb6exe)
        {
            if (string.IsNullOrWhiteSpace(vb6exe))
                throw new ArgumentException("Invalid VB6 executable path.", nameof(vb6exe));
            if (File.Exists(vb6exe) == false)
                throw new FileNotFoundException("Missing VB6 executable.");

            VB6Exe = vb6exe;
        }

        /// <summary>
        /// Path to the VB6 EXE.
        /// </summary>
        public string VB6Exe { get; }

        /// <summary>
        /// Preserves temporary files generated during the compilation.
        /// </summary>
        public bool PreserveTemporary { get; set; }

        /// <summary>
        /// Begins compilation. Returns <c>true</c> for successful compilation.
        /// </summary>
        /// <param name="vb6"></param>
        /// <param name="project"></param>
        /// <param name="output"></param>
        /// <param name="errors"></param>
        /// <returns></returns>
        public bool Compile(VB6Project project, string output, out IList<string> errors)
        {
            if (project == null)
                throw new ArgumentNullException(nameof(project));
            if (output == null)
                throw new ArgumentNullException(nameof(output));

            // create missing directory
            if (Directory.Exists(output) == false)
                Directory.CreateDirectory(output);

            var source = Path.Combine(Path.GetTempPath(), Path.ChangeExtension(Path.GetRandomFileName(), ".vbp"));

            try
            {
                project.Save(source);

                var sti = new ProcessStartInfo();
                sti.FileName = typeof(Executor).Assembly.Location;
                sti.Arguments = $@"-e ""{VB6Exe}"" -v ""{source}"" -o ""{output}""";
                sti.UseShellExecute = false;
                sti.CreateNoWindow = true;
                sti.RedirectStandardOutput = true;
                sti.RedirectStandardError = true;

                using (var prc = Process.Start(sti))
                {
                    var stderr = new List<string>();
                    prc.ErrorDataReceived += (s, a) => stderr.Add(a.Data);
                    prc.BeginErrorReadLine();
                    prc.BeginOutputReadLine();
                    prc.WaitForExit();

                    if (prc.ExitCode != 0)
                    {
                        errors = stderr.Select(i => i?.Trim()).Where(i => !string.IsNullOrWhiteSpace(i)).ToList();
                        return false;
                    }
                }
            }
            finally
            {
                try
                {
                    if (PreserveTemporary == false)
                        if (File.Exists(source))
                            File.Delete(source);
                }
                catch
                {

                }
            }

            errors = null;
            return true;
        }

    }

}
