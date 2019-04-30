using System.Collections.Generic;
using System;
using System.IO;
using System.Linq;
using System.Diagnostics;

using ILRepacking;

using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;

namespace Cogito.VisualBasic6.MSBuild
{

    public class VB6ILRepack :
        Microsoft.Build.Utilities.Task, IDisposable
    {

        string attributeFile;
        string logFile;
        string outputFile;
        string keyFile;
        string keyContainer;
        ITaskItem[] assemblies = new ITaskItem[0];
        ITaskItem[] libraryPath = new ITaskItem[0];
        ILRepack.Kind targetKind;
        bool parallel = true;
        ILRepack ilMerger;
        RepackOptions repackOptions;
        string excludeFileTmpPath;

        /// <summary>
        /// Specifies a keyfile to sign the output assembly.
        /// </summary>
        public virtual string KeyFile
        {
            get { return keyFile; }
            set { keyFile = BuildPath(ConvertEmptyToNull(value)); }
        }

        /// <summary>
        /// Specifies a KeyContainer to use.
        /// </summary>
        public virtual string KeyContainer
        {
            get { return keyContainer; }
            set { keyContainer = BuildPath(ConvertEmptyToNull(value)); }
        }

        /// <summary>
        /// Specifies a logfile to output log information.
        /// </summary>
        public virtual string LogFile
        {
            get { return logFile; }
            set { logFile = BuildPath(ConvertEmptyToNull(value)); }
        }

        /// <summary>
        /// Merges types with identical names into one.
        /// </summary>
        public virtual bool Union { get; set; }

        /// <summary>
        /// Enable/disable symbol file generation.
        /// </summary>
        public virtual bool DebugInfo { get; set; }

        /// <summary>
        /// Take assembly attributes from the given assembly file.
        /// </summary>
        public virtual string AttributeFile
        {
            get { return attributeFile; }
            set { attributeFile = BuildPath(ConvertEmptyToNull(value)); }
        }

        /// <summary>
        /// Copy assembly attributes (by default only the primary assembly attributes are copied).
        /// </summary>
        public virtual bool CopyAttributes { get; set; }

        /// <summary>
        /// Allows multiple attributes (if type allows).
        /// </summary>
        public virtual bool AllowMultiple { get; set; }

        /// <summary>
        /// Target assembly kind (Exe|Dll|WinExe|SameAsPrimaryAssembly).
        /// </summary>
        public virtual string TargetKind
        {
            get
            {
                return targetKind.ToString();
            }
            set
            {
                if (Enum.IsDefined(typeof(ILRepack.Kind), value))
                {
                    targetKind = (ILRepack.Kind)Enum.Parse(typeof(ILRepacking.ILRepack.Kind), value);
                }
                else
                {
                    Log.LogWarning("TargetKind should be [Exe|Dll|WinExe|SameAsPrimaryAssembly]; set to SameAsPrimaryAssembly");
                    targetKind = ILRepack.Kind.SameAsPrimaryAssembly;
                }
            }
        }

        /// <summary>
        /// Target platform (v1, v1.1, v2, v4 supported).
        /// </summary>
        public virtual string TargetPlatformVersion { get; set; }

        /// <summary>
        /// Path of Directory where target platform is located.
        /// </summary>
        public virtual string TargetPlatformDirectory { get; set; }

        /// <summary>
        /// Merge assembly xml documentation.
        /// </summary>
        public bool XmlDocumentation { get; set; }

        /// <summary>
        /// List of paths to use as "include directories" when attempting to merge assemblies.
        /// </summary>
        public virtual ITaskItem[] LibraryPath
        {
            get { return libraryPath; }
            set { libraryPath = value; }
        }

        /// <summary>
        /// Set all types but the ones from the first assembly 'internal'.
        /// </summary>
        public virtual bool Internalize { get; set; }

        /// <summary>
        /// Rename all internalized types (to be used when Internalize is enabled).
        /// </summary>
        public virtual bool RenameInternalized { get; set; }

        /// <summary>
        /// List of assemblies that should not be internalized.
        /// </summary>
        public virtual ITaskItem[] InternalizeExclude { get; set; }

        /// <summary>
        /// Output name for merged assembly.
        /// </summary>
        [Required]
        public virtual string OutputFile
        {
            get { return outputFile; }
            set
            {
                outputFile = ConvertEmptyToNull(value);
            }
        }

        /// <summary>
        /// List of assemblies that will be merged.
        /// </summary>
        [Required]
        public virtual ITaskItem[] InputAssemblies
        {
            get { return assemblies; }
            set { assemblies = value; }
        }

        /// <summary>
        /// Set the keyfile, but don't sign the assembly.
        /// </summary>
        public virtual bool DelaySign { get; set; }

        /// <summary>
        /// Allows to duplicate resources in output assembly (by default they're ignored).
        /// </summary>
        public virtual bool AllowDuplicateResources { get; set; }

        /// <summary>
        /// Allows assemblies with Zero PeKind (but obviously only IL will get merged).
        /// </summary>
        public virtual bool ZeroPeKind { get; set; }

        /// <summary>
        /// Use as many CPUs as possible to merge the assemblies.
        /// </summary>
        public virtual bool Parallel
        {
            get { return parallel; }
            set { parallel = value; }
        }

        /// <summary>
        /// Pause execution once completed (good for debugging).
        /// </summary>
        public virtual bool PauseBeforeExit { get; set; }

        /// <summary>
        /// Additional debug information during merge that will be outputted to LogFile.
        /// </summary>
        public virtual bool Verbose { get; set; }

        /// <summary>
        /// Allows (and resolves) file wildcards (e.g. `*`.dll) in input assemblies.
        /// </summary>
        public virtual bool Wildcards { get; set; }

        /// <summary>
        ///     Executes ILRepack with specified options.
        /// </summary>
        /// <returns>Returns true if its successful.</returns>
        public override bool Execute()
        {
            repackOptions = new RepackOptions
            {
                KeyFile = keyFile,
                KeyContainer = keyContainer,
                LogFile = logFile,
                Log = !string.IsNullOrEmpty(logFile),
                LogVerbose = Verbose,
                UnionMerge = Union,
                DebugInfo = DebugInfo,
                CopyAttributes = CopyAttributes,
                AttributeFile = AttributeFile,
                AllowMultipleAssemblyLevelAttributes = AllowMultiple,
                TargetKind = targetKind,
                TargetPlatformVersion = TargetPlatformVersion,
                TargetPlatformDirectory = TargetPlatformDirectory,
                XmlDocumentation = XmlDocumentation,
                Internalize = Internalize,
                RenameInternalized = RenameInternalized,
                DelaySign = DelaySign,
                AllowDuplicateResources = AllowDuplicateResources,
                AllowZeroPeKind = ZeroPeKind,
                Parallel = Parallel,
                PauseBeforeExit = PauseBeforeExit,
                OutputFile = outputFile,
                AllowWildCards = Wildcards
            };

            ilMerger = new ILRepack(repackOptions);

            // attempt to create output directory if it does not exist
            var outputPath = Path.GetDirectoryName(OutputFile);
            if (outputPath != null && Directory.Exists(outputPath) == false)
            {
                try
                {
                    Directory.CreateDirectory(outputPath);
                }
                catch (Exception ex)
                {
                    Log.LogErrorFromException(ex);
                    return false;
                }
            }

            // assemblies to be merged
            var assemblies = new string[this.assemblies.Length];
            for (var i = 0; i < this.assemblies.Length; i++)
            {
                assemblies[i] = this.assemblies[i].ItemSpec;

                if (string.IsNullOrEmpty(assemblies[i]))
                    throw new Exception($"Invalid assembly path on item index {i}");

                if (!File.Exists(assemblies[i]) && !File.Exists(BuildPath(assemblies[i])))
                    throw new Exception($"Unable to resolve assembly '{assemblies[i]}'");

                Log.LogMessage(MessageImportance.High, "Added assembly '{0}'", assemblies[i]);
            }

            // List of regex to compare against FullName of types NOT to internalize
            if (InternalizeExclude != null)
            {
                var internalizeExclude = new string[InternalizeExclude.Length];
                if (Internalize)
                {
                    for (var i = 0; i < InternalizeExclude.Length; i++)
                    {
                        internalizeExclude[i] = InternalizeExclude[i].ItemSpec;
                        if (string.IsNullOrEmpty(internalizeExclude[i]))
                        {
                            throw new Exception($"Invalid internalize exclude pattern at item index {i}. Pattern cannot be blank.");
                        }
                        Log.LogMessage(MessageImportance.High, "Excluding namespaces/types matching pattern '{0}' from being internalized", internalizeExclude[i]);
                    }

                    // Create a temporary file with a list of assemblies that should not be internalized.
                    excludeFileTmpPath = Path.GetTempFileName();
                    File.WriteAllLines(excludeFileTmpPath, internalizeExclude);
                    repackOptions.ExcludeFile = excludeFileTmpPath;
                }
            }

            repackOptions.InputAssemblies = assemblies;

            // Path that will be used when searching for assemblies to merge.
            var searchPath = new List<string> { "." };
            searchPath.AddRange(LibraryPath.Select(iti => BuildPath(iti.ItemSpec)));
            repackOptions.SearchDirectories = searchPath.ToArray();

            // Attempt to merge assemblies.
            try
            {
                Log.LogMessage(MessageImportance.High, "Merging {0} assemb{1} to '{2}'", this.assemblies.Length, this.assemblies.Length != 1 ? "ies" : "y", outputFile);

                // Measure performance
                var stopWatch = new Stopwatch();
                stopWatch.Start();
                ilMerger.Repack();
                stopWatch.Stop();

                Log.LogMessage(MessageImportance.High, "Merge succeeded in {0} s", stopWatch.Elapsed.TotalSeconds);
            }
            catch (Exception e)
            {
                Log.LogErrorFromException(e);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Converts empty string to null.
        /// </summary>
        /// <param name="str">String to check for emptiness</param>
        /// <returns></returns>
        private static string ConvertEmptyToNull(string str)
        {
            return string.IsNullOrEmpty(str) ? null : str;
        }

        /// <summary>
        /// Returns path respective to current working directory.
        /// </summary>
        /// <param name="path">Relative path to current working directory</param>
        /// <returns></returns>
        private string BuildPath(string path)
        {
            var workDir = Directory.GetCurrentDirectory();
            return string.IsNullOrEmpty(path) ? null : Path.Combine(workDir, path);
        }

        public void Dispose()
        {
            if (File.Exists(excludeFileTmpPath))
                File.Delete(excludeFileTmpPath);
        }

    }

}
