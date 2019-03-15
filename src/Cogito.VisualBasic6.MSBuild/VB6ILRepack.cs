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

        string _attributeFile;
        string _logFile;
        string _outputFile;
        string _keyFile;
        string _keyContainer;
        ITaskItem[] _assemblies = new ITaskItem[0];
        ITaskItem[] _libraryPath = new ITaskItem[0];
        ILRepacking.ILRepack.Kind _targetKind;
        bool _parallel = true;
        ILRepacking.ILRepack _ilMerger;
        RepackOptions _repackOptions;
        string _excludeFileTmpPath;

        /// <summary>
        /// Specifies a keyfile to sign the output assembly.
        /// </summary>
        public virtual string KeyFile
        {
            get { return _keyFile; }
            set { _keyFile = BuildPath(ConvertEmptyToNull(value)); }
        }

        /// <summary>
        /// Specifies a KeyContainer to use.
        /// </summary>
        public virtual string KeyContainer
        {
            get { return _keyContainer; }
            set { _keyContainer = BuildPath(ConvertEmptyToNull(value)); }
        }

        /// <summary>
        /// Specifies a logfile to output log information.
        /// </summary>
        public virtual string LogFile
        {
            get { return _logFile; }
            set { _logFile = BuildPath(ConvertEmptyToNull(value)); }
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
            get { return _attributeFile; }
            set { _attributeFile = BuildPath(ConvertEmptyToNull(value)); }
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
                return _targetKind.ToString();
            }
            set
            {
                if (Enum.IsDefined(typeof(ILRepacking.ILRepack.Kind), value))
                {
                    _targetKind = (ILRepacking.ILRepack.Kind)Enum.Parse(typeof(ILRepacking.ILRepack.Kind), value);
                }
                else
                {
                    Log.LogWarning("TargetKind should be [Exe|Dll|" +
                                   "WinExe|SameAsPrimaryAssembly]; " +
                                   "set to SameAsPrimaryAssembly");
                    _targetKind = ILRepacking.ILRepack.Kind.SameAsPrimaryAssembly;
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
            get { return _libraryPath; }
            set { _libraryPath = value; }
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
            get { return _outputFile; }
            set
            {
                _outputFile = ConvertEmptyToNull(value);
            }
        }

        /// <summary>
        /// List of assemblies that will be merged.
        /// </summary>
        [Required]
        public virtual ITaskItem[] InputAssemblies
        {
            get { return _assemblies; }
            set { _assemblies = value; }
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
            get { return _parallel; }
            set { _parallel = value; }
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
            _repackOptions = new RepackOptions
            {
                KeyFile = _keyFile,
                KeyContainer = _keyContainer,
                LogFile = _logFile,
                Log = !string.IsNullOrEmpty(_logFile),
                LogVerbose = Verbose,
                UnionMerge = Union,
                DebugInfo = DebugInfo,
                CopyAttributes = CopyAttributes,
                AttributeFile = AttributeFile,
                AllowMultipleAssemblyLevelAttributes = AllowMultiple,
                TargetKind = _targetKind,
                TargetPlatformVersion = TargetPlatformVersion,
                TargetPlatformDirectory = TargetPlatformDirectory,
                XmlDocumentation = XmlDocumentation,
                Internalize = Internalize,
                //RenameInternalized = RenameInternalized,
                DelaySign = DelaySign,
                AllowDuplicateResources = AllowDuplicateResources,
                AllowZeroPeKind = ZeroPeKind,
                Parallel = Parallel,
                PauseBeforeExit = PauseBeforeExit,
                OutputFile = _outputFile,
                AllowWildCards = Wildcards
            };
            _ilMerger = new ILRepacking.ILRepack(_repackOptions);

            // Attempt to create output directory if it does not exist.
            var outputPath = Path.GetDirectoryName(OutputFile);
            if (outputPath != null && !Directory.Exists(outputPath))
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

            // Assemblies to be merged.
            var assemblies = new string[_assemblies.Length];
            for (int i = 0; i < _assemblies.Length; i++)
            {
                assemblies[i] = _assemblies[i].ItemSpec;
                if (string.IsNullOrEmpty(assemblies[i]))
                {
                    throw new Exception(string.Format("Invalid assembly path on item index {0}", i));
                }
                if (!File.Exists(assemblies[i]) && !File.Exists(BuildPath(assemblies[i])))
                {
                    throw new Exception(string.Format("Unable to resolve assembly '{0}'", assemblies[i]));
                }
                Log.LogMessage(MessageImportance.High, "Added assembly '{0}'", assemblies[i]);
            }

            // List of regex to compare against FullName of types NOT to internalize
            if (InternalizeExclude != null)
            {
                var internalizeExclude = new string[InternalizeExclude.Length];
                if (Internalize)
                {
                    for (int i = 0; i < InternalizeExclude.Length; i++)
                    {
                        internalizeExclude[i] = InternalizeExclude[i].ItemSpec;
                        if (string.IsNullOrEmpty(internalizeExclude[i]))
                        {
                            throw new Exception(string.Format("Invalid internalize exclude pattern at item index {0}. Pattern cannot be blank.", i));
                        }
                        Log.LogMessage(MessageImportance.High,
                            "Excluding namespaces/types matching pattern '{0}' from being internalized", internalizeExclude[i]);
                    }

                    // Create a temporary file with a list of assemblies that should not be internalized.
                    _excludeFileTmpPath = Path.GetTempFileName();
                    File.WriteAllLines(_excludeFileTmpPath, internalizeExclude);
                    _repackOptions.ExcludeFile = _excludeFileTmpPath;
                }
            }

            _repackOptions.InputAssemblies = assemblies;

            // Path that will be used when searching for assemblies to merge.
            var searchPath = new List<string> { "." };
            searchPath.AddRange(LibraryPath.Select(iti => BuildPath(iti.ItemSpec)));
            _repackOptions.SearchDirectories = searchPath.ToArray();

            // Attempt to merge assemblies.
            try
            {
                Log.LogMessage(MessageImportance.High, "Merging {0} assemb{1} to '{2}'",
                    _assemblies.Length, _assemblies.Length != 1 ? "ies" : "y", _outputFile);

                // Measure performance
                Stopwatch stopWatch = new Stopwatch();
                stopWatch.Start();

                _ilMerger.Repack();
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
            // Remove temporary exclude file
            if (File.Exists(_excludeFileTmpPath))
            {
                File.Delete(_excludeFileTmpPath);
            }
        }

    }

}
