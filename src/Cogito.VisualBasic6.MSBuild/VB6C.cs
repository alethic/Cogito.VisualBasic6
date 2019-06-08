using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using Cogito.VisualBasic6.VB6C;
using Cogito.VisualBasic6.VB6C.Project;

using Microsoft.Build.Framework;

namespace Cogito.VisualBasic6.MSBuild.Tasks
{

    public class VB6C :
        Microsoft.Build.Utilities.Task
    {

        VB6Type type;

        /// <summary>
        /// Path to the VB6 executable.
        /// </summary>
        [Required]
        public string ToolPath { get; set; }

        /// <summary>
        /// Source VBP project file to use as a base.
        /// </summary>
        public string Source { get; set; }

        /// <summary>
        /// Type of the project.
        /// </summary>
        [Required]
        public string Type
        {
            get => Enum.GetName(typeof(VB6Type), type);
            set => type = (VB6Type)Enum.Parse(typeof(VB6Type), value);
        }

        /// <summary>
        /// Type of startup.
        /// </summary>
        public string Startup { get; set; }

        /// <summary>
        /// References to include in the VBP.
        /// </summary>
        public ITaskItem[] References { get; set; }

        /// <summary>
        /// Modules to include in the VBP.
        /// </summary>
        public ITaskItem[] Modules { get; set; }

        /// <summary>
        /// Classes to include in the VBP.
        /// </summary>
        public ITaskItem[] Classes { get; set; }

        /// <summary>
        /// Forms to include in the VBP.
        /// </summary>
        public ITaskItem[] Forms { get; set; }

        /// <summary>
        /// Additional properties to apply to the VBP.
        /// </summary>
        public string Properties { get; set; }

        /// <summary>
        /// Path to the output directory.
        /// </summary>
        [Required]
        public string Output { get; set; }

        /// <summary>
        /// Whether temporary files should be preserved.
        /// </summary>
        public bool PreserveTemporary { get; set; }

        /// <summary>
        /// URI to an output cache.
        /// </summary>
        public string OutputCacheUri { get; set; }

        /// <summary>
        /// Provider name of the output cache.
        /// </summary>
        public string OutputCacheProvider { get; set; } = "file";

        /// <summary>
        /// Applies the transforms.
        /// </summary>
        /// <param name="project"></param>
        /// <returns></returns>
        void Apply(VB6Project project)
        {
            project.Type = type;
            project.Startup = Startup;

            if (References != null &&
                References.Length > 0)
            {
                project.References.Clear();
                project.References.AddRange(GetReferencesForTaskItems(References).Where(i => i != null));
            }

            if (Modules != null &&
                Modules.Length > 0)
            {
                project.Modules.Clear();
                project.Modules.AddRange(GetModulesForTaskItems(Modules));
            }

            if (Classes != null &&
                Classes.Length > 0)
            {
                project.Classes.Clear();
                project.Classes.AddRange(GetClassesForTaskItems(Classes));
            }

            if (Forms != null &&
                Forms.Length > 0)
            {
                project.Forms.Clear();
                project.Forms.AddRange(GetFormsForTaskItems(Forms));
            }

            if (!string.IsNullOrWhiteSpace(Properties))
            {
                foreach (var kvp in Properties
                    .Split(';')
                    .Select(i => i.Split('='))
                    .Where(i => i.Length == 2)
                    .Where(i => !string.IsNullOrWhiteSpace(i[0]))
                    .Where(i => !string.IsNullOrWhiteSpace(i[1])))
                {
                    var key = kvp[0].Trim();
                    var val = kvp[1].Trim();
                    var typ = GetPropertyValueType(project, key, val);
                    if (typ != null)
                        project.Properties[key] = Convert.ChangeType(val, typ);
                }
            }
        }

        /// <summary>
        /// Attempt to get the appropriate property value type for the given key.
        /// </summary>
        /// <param name="project"></param>
        /// <param name="key"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        Type GetPropertyValueType(VB6Project project, string key, string value)
        {
            // project might already have a known value
            if (project.Properties.ContainsKey(key))
            {
                var exist = project.Properties[key];
                if (exist != null)
                    return exist.GetType();
            }

            // convert to int
            if (int.TryParse(value, out _))
                return typeof(int);

            // default to string
            return typeof(string);
        }

        /// <summary>
        /// Gets the set of <see cref="VB6ReferenceItem"/> entries from the given task item set.
        /// </summary>
        /// <param name="items"></param>
        /// <returns></returns>
        IEnumerable<VB6ReferenceItem> GetReferencesForTaskItems(ITaskItem[] items)
        {
            foreach (var item in items)
                yield return CreateReferenceForTaskItem(item);
        }

        /// <summary>
        /// Converts a task item into a <see cref="VB6ReferenceItem"/>.
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        VB6ReferenceItem CreateReferenceForTaskItem(ITaskItem item)
        {
            if (string.IsNullOrWhiteSpace(item.GetMetadata("TypeLibPath")))
            {
                Log.LogWarning("{0}: Item contains no TypeLibPath.", item.ItemSpec);
                return null;
            }

            Log.LogMessage("TypeLibPath is {0}", item.GetMetadata("TypeLibPath"));

            return new VB6ReferenceItem(
                Guid.Parse(item.GetMetadata("Guid")),
                new Version(
                    int.Parse(item.GetMetadata("VersionMajor")),
                    int.Parse(item.GetMetadata("VersionMinor"))),
                int.Parse(item.GetMetadata("Lcid")),
                item.GetMetadata("TypeLibPath"),
                item.GetMetadata("Name"),
                item.GetMetadata("Description"));
        }

        /// <summary>
        /// Gets the set of <see cref="VB6ModuleItem"/> entries from the given task item set.
        /// </summary>
        /// <param name="items"></param>
        /// <returns></returns>
        IEnumerable<VB6ModuleItem> GetModulesForTaskItems(ITaskItem[] items)
        {
            foreach (var item in items)
                yield return CreateModuleForTaskItem(item);
        }

        /// <summary>
        /// Converts a task item into a <see cref="VB6ModuleItem"/>.
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        VB6ModuleItem CreateModuleForTaskItem(ITaskItem item)
        {
            return new VB6ModuleItem(
                GetModuleNameFromMetadata(item) ?? Path.GetFileNameWithoutExtension(item.ItemSpec),
                GetFullPath(item));
        }

        /// <summary>
        /// Extracts the module name from metadata if available.
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        string GetModuleNameFromMetadata(ITaskItem item)
        {
            return !string.IsNullOrWhiteSpace(item.GetMetadata("Name")) ? item.GetMetadata("Name") : null;
        }

        /// <summary>
        /// Gets the set of <see cref="VB6ClassItem"/> entries from the given task item set.
        /// </summary>
        /// <param name="items"></param>
        /// <returns></returns>
        IEnumerable<VB6ClassItem> GetClassesForTaskItems(ITaskItem[] items)
        {
            foreach (var item in items)
                yield return CreateClassForTaskItem(item);
        }

        /// <summary>
        /// Converts a task item into a <see cref="VB6ClassItem"/>.
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        VB6ClassItem CreateClassForTaskItem(ITaskItem item)
        {
            return new VB6ClassItem(
                GetClassNameFromMetadata(item) ?? Path.GetFileNameWithoutExtension(item.ItemSpec),
                GetFullPath(item));
        }

        /// <summary>
        /// Extracts the class name from metadata if available.
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        string GetClassNameFromMetadata(ITaskItem item)
        {
            return !string.IsNullOrWhiteSpace(item.GetMetadata("Name")) ? item.GetMetadata("Name") : null;
        }

        /// <summary>
        /// Gets the set of <see cref="VB6FormItem"/> entries from the given task item set.
        /// </summary>
        /// <param name="items"></param>
        /// <returns></returns>
        IEnumerable<VB6FormItem> GetFormsForTaskItems(ITaskItem[] items)
        {
            foreach (var item in items)
                yield return CreateFormForTaskItem(item);
        }

        /// <summary>
        /// Converts a task item into a <see cref="VB6FormItem"/>.
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        VB6FormItem CreateFormForTaskItem(ITaskItem item)
        {
            return new VB6FormItem(GetFullPath(item));
        }

        /// <summary>
        /// Makes a relative path to the item given the target project file.
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        string GetFullPath(ITaskItem item)
        {
            return !string.IsNullOrWhiteSpace(item.GetMetadata("FullPath")) ? item.GetMetadata("FullPath") : Path.GetFullPath(item.ItemSpec);
        }

        /// <summary>
        /// Executes the task.
        /// </summary>
        /// <returns></returns>
        public override bool Execute()
        {
            // setup VB6 project
            var log = new StringWriter();
            var source = Source != null ? VB6Project.Load(Source) : new VB6Project();
            Apply(source);

            // output to random directory to prevent it from deleting contents of original
            var dst = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());

            try
            {
                Directory.CreateDirectory(dst);

                using (var mutex = new Mutex(true, new Guid(MD5.Create().ComputeHash(Encoding.UTF8.GetBytes(Output))).ToString("n")))
                {
                    if (GetOutputCache(source, dst) == false)
                    {
                        var compiler = new Compiler(ToolPath)
                        {
                            PreserveTemporary = PreserveTemporary
                        };

                        if (compiler.Compile(source, dst, out var errors) == false)
                        {
                            foreach (var error in errors)
                                Log.LogError(error);

                            return false;
                        }

                        // attempt to store the output into a cache
                        StoreOutputCache(source, dst);
                    }

                    // copy temporary directory to final
                    DirectoryCopy(dst, Output);
                }

                // return output
                Log.LogMessagesFromStream(new StringReader(log.ToString()), MessageImportance.Normal);
                return true;
            }
            finally
            {
                try
                {
                    if (PreserveTemporary == false)
                        if (Directory.Exists(dst))
                            Directory.Delete(dst, true);
                }
                catch
                {

                }
            }
        }

        /// <summary>
        /// Attempts to retrieve the compiler output from the specified cache provider.
        /// </summary>
        /// <returns></returns>
        bool GetOutputCache(VB6Project source, string output)
        {
            if (OutputCacheProvider == null)
                return false;

            switch (OutputCacheProvider)
            {
                case "file":
                    return GetFileOutputCache(source, output);
                default:
                    return false;
            }
        }

        void StoreOutputCache(VB6Project source, string output)
        {
            if (OutputCacheProvider == null)
                return;

            switch (OutputCacheProvider)
            {
                case "file":
                    StoreFileOutputCache(source, output);
                    break;
            }
        }

        /// <summary>
        /// Returns a unique hash value for the specified VB6 project.
        /// </summary>
        /// <param name="project"></param>
        /// <returns></returns>
        byte[] HashVB6Project(VB6Project project)
        {
            using (var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA1))
            {
                hash.AppendData(BitConverter.GetBytes((int)project.Type));

                hash.AppendData(BitConverter.GetBytes(project.References.Count));

                foreach (var i in project.References.OrderBy(i => i.Guid))
                {
                    hash.AppendData(i.Guid.ToByteArray());
                    hash.AppendData(BitConverter.GetBytes(i.Version.Major));
                    hash.AppendData(BitConverter.GetBytes(i.Version.Minor));
                    hash.AppendData(BitConverter.GetBytes(i.Version.Build));
                    hash.AppendData(BitConverter.GetBytes(i.Version.Revision));
                    hash.AppendData(BitConverter.GetBytes(i.LCID));

                    if (File.Exists(i.Location))
                        hash.AppendData(File.ReadAllBytes(i.Location));
                }

                hash.AppendData(BitConverter.GetBytes(project.Modules.Count));

                foreach (var i in project.Modules.OrderBy(i => i.Name))
                {
                    hash.AppendData(Encoding.UTF8.GetBytes(i.Name));

                    if (File.Exists(i.File))
                        hash.AppendData(File.ReadAllBytes(i.File));
                }

                hash.AppendData(BitConverter.GetBytes(project.Classes.Count));

                foreach (var i in project.Classes.OrderBy(i => i.Name))
                {
                    hash.AppendData(Encoding.UTF8.GetBytes(i.Name));

                    if (File.Exists(i.File))
                        hash.AppendData(File.ReadAllBytes(i.File));
                }

                hash.AppendData(BitConverter.GetBytes(project.Forms.Count));

                foreach (var i in project.Forms.OrderBy(i => i.File))
                {
                    if (File.Exists(i.File))
                        hash.AppendData(File.ReadAllBytes(i.File));
                }

                return hash.GetHashAndReset();
            }
        }

        /// <summary>
        /// Returns a unique string representing the project.
        /// </summary>
        /// <param name="project"></param>
        /// <returns></returns>
        string HashVB6ProjectToString(VB6Project project)
        {
            return string.Join("", HashVB6Project(project).Select(i => i.ToString("x2")));
        }

        /// <summary>
        /// Attempts to retrieve the VB6 project output from the output cache.
        /// </summary>
        /// <param name="project"></param>
        /// <param name="output"></param>
        /// <returns></returns>
        bool GetFileOutputCache(VB6Project project, string output)
        {
            if (!Uri.TryCreate(OutputCacheUri, UriKind.Absolute, out var uri) || uri.Scheme != "file")
                throw new InvalidOperationException($"Bad file cache URI: {OutputCacheUri}.");

            Log.LogMessage(MessageImportance.High, "Retrieving VB6 output from file cache: {0}", uri);

            var hash = HashVB6ProjectToString(project);
            var path = Path.Combine(uri.LocalPath, hash);
            if (Directory.Exists(path))
            {
                DirectoryCopy(path, output);
                return true;
            }

            return false;
        }

        /// <summary>
        /// Attempts to store the VB6 project output into the file cache.
        /// </summary>
        /// <param name="project"></param>
        /// <param name="source"></param>
        void StoreFileOutputCache(VB6Project project, string source)
        {
            if (!Uri.TryCreate(OutputCacheUri, UriKind.Absolute, out var uri) || uri.Scheme != "file")
                throw new InvalidOperationException($"Bad file cache URI: {OutputCacheUri}.");

            Log.LogMessage(MessageImportance.High, "Storing VB6 output to file cache: {0}", uri);

            var hash = HashVB6ProjectToString(project);
            var path = Path.Combine(uri.LocalPath, hash);
            if (Directory.Exists(path) == false)
                Directory.CreateDirectory(path);

            DirectoryCopy(source, path);
        }

        /// <summary>
        /// Copies the contents of one directory to another.
        /// </summary>
        /// <param name="sourceDirName"></param>
        /// <param name="destDirName"></param>
        /// <param name="copySubDirs"></param>
        static void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs = true)
        {
            if (sourceDirName == null)
                throw new ArgumentNullException(nameof(sourceDirName));
            if (destDirName == null)
                throw new ArgumentNullException(nameof(destDirName));

            var dir = new DirectoryInfo(sourceDirName);
            if (dir.Exists == false)
                throw new DirectoryNotFoundException($"Source directory does not exist or could not be found: {sourceDirName}");

            if (Directory.Exists(destDirName) == false)
                Directory.CreateDirectory(destDirName);

            foreach (var file in dir.GetFiles())
                file.CopyTo(Path.Combine(destDirName, file.Name), true);

            if (copySubDirs)
                foreach (var subdir in dir.GetDirectories())
                    DirectoryCopy(subdir.FullName, Path.Combine(destDirName, subdir.Name), copySubDirs);
        }

    }

}
