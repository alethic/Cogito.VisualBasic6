using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.Remoting;
using EasyHook;

namespace Cogito.VisualBasic6.Make
{

    /// <summary>
    /// Allows for execution of a wrapped VB6 instance.
    /// </summary>
    public class Executor : MarshalByRefObject, IDisposable
    {

        /// <summary>
        /// Temporary path to a log file.
        /// </summary>
        readonly string log = Path.GetTempFileName();

        /// <summary>
        /// Path to the original EXE.
        /// </summary>
        public string Exe { get; set; }

        /// <summary>
        /// Path to the source project file to be built.
        /// </summary>
        public string Vbp { get; set; }

        /// <summary>
        /// Path to the output file to be built.
        /// </summary>
        public string Out { get; set; }

        /// <summary>
        /// Path to the output directory.
        /// </summary>
        public string Dir { get; set; }

        /// <summary>
        /// Additional defines.
        /// </summary>
        public Dictionary<string, string> Def { get; set; }

        /// <summary>
        /// Prepares for execution.
        /// </summary>
        void Prepare()
        {
            if (string.IsNullOrWhiteSpace(Exe))
                throw new InvalidOperationException("Invalid VB6 executable.");
            if (string.IsNullOrWhiteSpace(Vbp))
                throw new InvalidOperationException("Invalid VB6 project file.");
            if (string.IsNullOrWhiteSpace(Out))
                throw new InvalidOperationException("Invalid output path.");

            Exe = Exe?.Trim().TrimEnd(new[] { '/', '\\' });
            Vbp = Vbp?.Trim().TrimEnd(new[] { '/', '\\' });
            Out = Out?.Trim().TrimEnd(new[] { '/', '\\' });

            if (File.Exists(Exe) == false)
                throw new FileNotFoundException("Missing VB6 executable.", Exe);
            if (File.Exists(Vbp) == false)
                throw new FileNotFoundException("Missing VB6 project file.", Vbp);

            if (string.IsNullOrWhiteSpace(Dir))
                Dir = Environment.CurrentDirectory;

            Dir = Dir.Trim().TrimEnd(new[] { '/', '\\' });

            if (Def == null)
                Def = new Dictionary<string, string>();
        }

        /// <summary>
        /// Invoked periodically by VB6.
        /// </summary>
        public void Ping()
        {

        }

        /// <summary>
        /// Invoked to write a value to stderr.
        /// </summary>
        /// <param name="value"></param>
        public void WriteStdErr(string value)
        {
            if (value != null)
                Console.Error.Write(value);
        }

        /// <summary>
        /// Invoked to write a value to stderr.
        /// </summary>
        /// <param name="value"></param>
        public void WriteStdOut(string value)
        {
            if (value != null)
                Console.Write(value);
        }

        /// <summary>
        /// Executes VB6.
        /// </summary>
        public void Execute()
        {
            Prepare();

            if (File.Exists(log))
                File.Delete(log);

            // generate new channel to call back into parent
            string channelName = null;
            RemoteHooking.IpcCreateServer(ref channelName, WellKnownObjectMode.Singleton, this);

            // spawn VB6 process
            RemoteHooking.CreateAndInject(
                Exe,
                BuildArgs(),
                0,
                InjectionOptions.DoNotRequireStrongName,
                System.Reflection.Assembly.GetExecutingAssembly().Location,
                System.Reflection.Assembly.GetExecutingAssembly().Location,
                out var pid,
                channelName);

            // wait for exit of process
            var prc = Process.GetProcessById(pid);
            while (prc != null && !prc.HasExited)
                prc.WaitForExit(100);

            // copy output to stderr
            if (File.Exists(log))
                using (var l = new StreamReader(File.OpenRead(log)))
                    while (l.ReadLine() is string s)
                        Console.Error.WriteLine(s);
        }

        /// <summary>
        /// Builds the command arguments to be executed.
        /// </summary>
        /// <returns></returns>
        string BuildArgs()
        {
            var l = new List<string>();

            l.Add("/m");
            l.Add('"' + Vbp + '"');

            l.Add("/out");
            l.Add('"' + log + '"');

            if (!string.IsNullOrWhiteSpace(Dir))
            {
                l.Add("/outdir");
                l.Add('"' + Dir + '"');
            }

            if (Def.Count > 0)
            {
                l.Add("/D");
                l.Add(string.Join(":", Def.Select(i => $"{i.Key}={i.Value}")));
            }

            if (!string.IsNullOrWhiteSpace(Out))
                l.Add('"' + Out + '"');

            return string.Join(" ", l);
        }

        /// <summary>
        /// Disposes of the instance.
        /// </summary>
        public void Dispose()
        {
            try
            {
                if (File.Exists(log))
                    File.Delete(log);
            }
            catch
            {

            }

            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Finalizes the instance.
        /// </summary>
        ~Executor()
        {
            Dispose();
        }

    }

}
