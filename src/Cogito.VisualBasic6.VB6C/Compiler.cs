using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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
        /// Compiles a <see cref="VB6Project"/>.
        /// </summary>
        /// <param name="vb6"></param>
        /// <param name="project"></param>
        /// <param name="output"></param>
        /// <param name="logger"></param>
        public void Compile(FileInfo vb6, VB6Project project, DirectoryInfo output, TextWriter logger)
        {
            if (vb6 == null)
                throw new ArgumentNullException(nameof(vb6));
            if (project == null)
                throw new ArgumentNullException(nameof(project));
            if (output == null)
                throw new ArgumentNullException(nameof(output));

            var source = Path.Combine(Path.GetTempPath(), Path.ChangeExtension(Path.GetRandomFileName(), ".vbp"));

            try
            {
                project.Save(source);

                var sti = new ProcessStartInfo(typeof(Executor).Assembly.Location, $@"-e ""{vb6}"" -v ""{source}"" -o ""{output.FullName}""");
                sti.UseShellExecute = false;
                sti.RedirectStandardOutput = true;
                sti.RedirectStandardError = true;
                using (var prc = Process.Start(sti))
                {
                    prc.OutputDataReceived += (s, a) => logger.WriteLine(a.Data);
                    prc.ErrorDataReceived += (s, a) => logger.WriteLine(a.Data);
                    prc.BeginErrorReadLine();
                    prc.BeginOutputReadLine();
                    prc.WaitForExit();
                }
            }
            finally
            {
                try
                {
                    if (File.Exists(source))
                        File.Delete(source);
                }
                catch
                {

                }
            }
        }

    }

}
