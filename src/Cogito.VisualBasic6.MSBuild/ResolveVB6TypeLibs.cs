using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;
using Microsoft.Win32;

namespace Cogito.VisualBasic6.MSBuild
{

    public class ResolveVB6TypeLibs :
        Microsoft.Build.Utilities.Task
    {

        class TypeLibInfo
        {

            public Guid Guid { get; set; }

            public int? MajorVersion { get; set; }

            public int? MinorVersion { get; set; }

            public int? Lcid { get; set; }

            public string Name { get; set; }

            public string Description { get; set; }

            public string TypeLibPath { get; set; }

            public string TypeLibFilePath { get; set; }

        }

        public string WorkingDirectory { get; set; }

        /// <summary>
        /// References to include in the VBP.
        /// </summary>
        public ITaskItem[] COMReference { get; set; }

        /// <summary>
        /// Modules to include in the VBP.
        /// </summary>
        [Output]
        public ITaskItem[] TypeLibItems { get; set; }

        /// <summary>
        /// Returns <c>null</c> for an empty string.
        /// </summary>
        /// <param name="t"></param>
        /// <returns></returns>
        string TrimToNull(string t)
        {
            return !string.IsNullOrWhiteSpace(t) ? t.Trim() : null;
        }

        /// <summary>
        /// Attempts to convert a relative path to an absolute path.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        string NormalizePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return null;

            // make rooted from current project directory
            if (Path.IsPathRooted(path) == false)
                path = Path.Combine(WorkingDirectory, path);

            // normalize results
            return Path.GetFullPath(new Uri(path).LocalPath)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        }

        [DllImport("oleaut32.dll", PreserveSig = false)]
        public static extern ITypeLib LoadTypeLib([In, MarshalAs(UnmanagedType.LPWStr)] string typelib);

        readonly Dictionary<string, Tuple<bool, TypeLibInfo>> typeLibCache = new Dictionary<string, Tuple<bool, TypeLibInfo>>();

        /// <summary>
        /// Attempts to load the specified type lib and return it's parsed metadata.
        /// </summary>
        /// <param name="path"></param>
        /// <param name="info"></param>
        /// <returns></returns>
        bool TryLoadTypeLibFromPath(string path, out TypeLibInfo info)
        {
            info = null;

            if (string.IsNullOrWhiteSpace(path))
                return false;

            // try from cache
            if (typeLibCache.TryGetValue(path, out var cached))
            {
                info = cached.Item2;
                return cached.Item1;
            }

            // attempt to load actual path
            var b = TryLoadTypeLibFromPathInternal(path, out info);
            typeLibCache[path] = Tuple.Create(b, info);
            return b;
        }

        /// <summary>
        /// Attempts to advance the path upwards until a real file is detected. Escapes from the /3 notation used by
        /// some HKCR entries.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        string GetTypeLibFilePath(string path)
        {
            while (!File.Exists(path))
                path = Path.GetDirectoryName(path);

            return path;
        }

        /// <summary>
        /// Attempts to load the specified type lib and return it's parsed metadata.
        /// </summary>
        /// <param name="typeLibPath"></param>
        /// <returns></returns>
        bool TryLoadTypeLibFromPathInternal(string path, out TypeLibInfo info)
        {
            info = null;

            if (string.IsNullOrWhiteSpace(path))
                return false;

            Log.LogMessage("TryLoadTypeLibFromPath: {0}", path);

            ITypeLib typeLib = null;
            var typeLibPtr = IntPtr.Zero;

            try
            {
                typeLib = LoadTypeLib(path);
                if (typeLib != null)
                {
                    typeLib.GetLibAttr(out typeLibPtr);
                    if (typeLibPtr != IntPtr.Zero)
                    {
                        // marshal pointer into struct
                        var ta = (System.Runtime.InteropServices.ComTypes.TYPELIBATTR)Marshal.PtrToStructure(typeLibPtr, typeof(System.Runtime.InteropServices.ComTypes.TYPELIBATTR));

                        // obtain information from typelib
                        typeLib.GetDocumentation(-1, out var name, out var docString, out var helpContext, out var helpFile);

                        // generate information
                        info = new TypeLibInfo()
                        {
                            Name = name,
                            Description = docString,
                            Guid = ta.guid,
                            MajorVersion = ta.wMajorVerNum,
                            MinorVersion = ta.wMinorVerNum,
                            Lcid = ta.lcid,
                            TypeLibPath = NormalizePath(path),
                            TypeLibFilePath = GetTypeLibFilePath(path),
                        };

                        // success
                        return true;
                    }
                }
            }
            catch (COMException e)
            {
                Log.LogMessage(MessageImportance.Low, "LoadTypeLib error: {0}; {1}", path, e.Message);
            }
            finally
            {
                if (typeLib != null && typeLibPtr != IntPtr.Zero)
                    typeLib.ReleaseTLibAttr(typeLibPtr);
            }

            return false;
        }

        /// <summary>
        /// Probes for the type lib path for the given COM class.
        /// </summary>
        /// <param name="classId"></param>
        /// <param name="majorVersion"></param>
        /// <param name="minorVersion"></param>
        /// <param name="lcid"></param>
        /// <returns></returns>
        string GetTypeLibForClass(Guid classId, int majorVersion, int minorVersion, int lcid)
        {
            var b = new StringBuilder();
            b.Append("HKEY_CLASSES_ROOT").Append('\\');
            b.Append("TypeLib").Append('\\');
            b.Append("{").Append(classId).Append("}").Append('\\');
            b.Append(majorVersion).Append(".").Append(minorVersion).Append('\\');
            b.Append(lcid).Append('\\');
            b.Append("win32");
            var p = NormalizePath((string)Registry.GetValue(b.ToString(), null, null));
            return p;
        }

        /// <summary>
        /// Attempts to parse the given <see cref="int"/>.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        int? ParseInt(string value)
        {
            if (value != null)
                if (int.TryParse(value, out var i))
                    return i;

            return null;
        }

        /// <summary>
        /// Extracts the type lib paths on the given item.
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        IEnumerable<string> GetTypeLibPathMetadata(ITaskItem item)
        {
            foreach (var i in item.GetMetadata("TypeLibPath").Split(';'))
                yield return NormalizePath(TrimToNull(i));
        }

        /// <summary>
        /// Returns the type lib information for the given task item.
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        TypeLibInfo GetTypeLibInfo(ITaskItem item)
        {
            TypeLibInfo info;

            // type lib path specified, load information
            foreach (var typeLibPath in GetTypeLibPathMetadata(item))
                if (TryLoadTypeLibFromPath(typeLibPath, out info))
                    return info;

            // full path specified, probe for type lib
            if (TryLoadTypeLibFromPath(NormalizePath(TrimToNull(item.GetMetadata("FullPath"))), out info))
                return info;

            // identity might lead to a file
            var identity = TrimToNull(item.ItemSpec);
            if (TryLoadTypeLibFromPath(NormalizePath(identity), out info))
                return info;

            // guid specified, result to registry probe
            var guid = !string.IsNullOrWhiteSpace(item.GetMetadata("Guid")) ? (Guid?)Guid.Parse(TrimToNull(item.GetMetadata("Guid"))) : null;
            var majorVersion = ParseInt(TrimToNull(item.GetMetadata("VersionMajor")));
            var minorVersion = ParseInt(TrimToNull(item.GetMetadata("VersionMinor")));
            if (guid != null && majorVersion != null && minorVersion != null)
                if (TryLoadTypeLibFromPath(GetTypeLibForClass((Guid)guid, (int)majorVersion, (int)minorVersion, ParseInt(TrimToNull(item.GetMetadata("Lcid"))) ?? 0), out info))
                    return info;

            return null;
        }

        /// <summary>
        /// Attempts to resolve the given type lib item with additional metadata.
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        ITaskItem ResolveTypeLibItem(ITaskItem item)
        {
            // acquire type library information
            var info = GetTypeLibInfo(item);
            if (info == null)
                return null;

            Log.LogMessage(MessageImportance.High, "ResolveTypeLib: {0} -> {1}", item.ItemSpec, info?.TypeLibPath);

            // generate new task item for resulting type lib information
            return new TaskItem(item.ItemSpec ?? Path.GetFileName(info.TypeLibPath), new Dictionary<string, string>()
            {
                ["Name"] = info.Name,
                ["Description"] = info.Description,
                ["Guid"] = info.Guid.ToString(),
                ["VersionMajor"] = info.MajorVersion.ToString(),
                ["VersionMinor"] = info.MinorVersion.ToString(),
                ["Lcid"] = info.Lcid.ToString(),
                ["TypeLibPath"] = info.TypeLibPath,
                ["TypeLibFilePath"] = info.TypeLibFilePath
            });
        }

        /// <summary>
        /// Filters out duplicate type lib references.
        /// </summary>
        /// <param name="items"></param>
        /// <returns></returns>
        IEnumerable<ITaskItem> DistinctTypeLib(IEnumerable<ITaskItem> items)
        {
            var hs = new HashSet<string>();
            foreach (var i in items)
                if (hs.Add(i.GetMetadata("Guid")))
                    yield return i;
        }

        /// <summary>
        /// Executes the task.
        /// </summary>
        /// <returns></returns>
        public override bool Execute()
        {
            TypeLibItems = DistinctTypeLib(COMReference.Select(i => ResolveTypeLibItem(i)).Where(i => i != null)).OrderBy(i => i.GetMetadata("TypeLibFilePath")).ToArray();
            return true;
        }

    }



}
