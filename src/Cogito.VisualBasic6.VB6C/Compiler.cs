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
        /// Compiles a <see cref="VB6Project"/>.
        /// </summary>
        /// <param name="vb6"></param>
        /// <param name="project"></param>
        /// <param name="output"></param>
        /// <param name="logger"></param>
        public List<string> Compile(FileInfo vb6, VB6Project project, DirectoryInfo output)
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

                var sti = new ProcessStartInfo();
                sti.FileName = typeof(Executor).Assembly.Location;
                sti.Arguments = $@"-e ""{vb6}"" -v ""{source}"" -o ""{output.FullName}""";
                sti.UseShellExecute = false;
                sti.RedirectStandardOutput = true;
                sti.RedirectStandardError = true;

                using (var prc = Process.Start(sti))
                {
                    var errors = new List<string>();
                    prc.ErrorDataReceived += (s, a) => errors.Add(a.Data);
                    prc.BeginErrorReadLine();
                    prc.BeginOutputReadLine();
                    prc.WaitForExit();
                    if (prc.ExitCode != 0)
                        return errors.Select(i => i?.Trim()).Where(i => !string.IsNullOrWhiteSpace(i)).ToList();
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

            return null;
        }

    }

}
