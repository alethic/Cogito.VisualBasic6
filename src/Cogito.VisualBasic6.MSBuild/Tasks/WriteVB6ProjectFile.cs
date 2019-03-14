using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Microsoft.Build.Framework;

namespace Cogito.VisualBasic6.MSBuild
{

    public class WriteVB6ProjectFile :
        Microsoft.Build.Utilities.Task
    {

        enum VB6Type
        {

            Exe,
            OleDll,

        }


        class VB6ReferenceItem
        {

            static Regex REGEX = new Regex(@"\*\\G(?<g>\{.+\})#(?<v>\d\.\d)#(?<c>\d+)#(?<p>.+)#(?<h>.+)");

            /// <summary>
            /// Parses the given reference definition.
            /// </summary>
            /// <param name="definition"></param>
            /// <returns></returns>
            public static VB6ReferenceItem Parse(string definition)
            {
                var m = REGEX.Match(definition);
                if (m.Success)
                {
                    var g = Guid.Parse(m.Groups["g"].Value);
                    var v = Version.Parse(m.Groups["v"].Value);
                    var c = int.Parse(m.Groups["c"].Value);
                    var p = m.Groups["p"].Value;
                    var h = m.Groups["h"].Value;
                    return new VB6ReferenceItem(g, v, c, p, null, h);
                }

                throw new FormatException("Cannot parse Reference: '" + definition + "'");
            }

            /// <summary>
            /// Initializes a new instance.
            /// </summary>
            /// <param name="guid"></param>
            /// <param name="version"></param>
            /// <param name="lcid"></param>
            /// <param name="location"></param>
            /// <param name="name"></param>
            /// <param name="description"></param>
            public VB6ReferenceItem(Guid guid, Version version, int lcid, string location, string name, string description)
            {
                Guid = guid;
                Version = version;
                LCID = lcid;
                Location = location;
                Name = name;
                Description = description;
            }

            public Guid Guid { get; set; }

            public Version Version { get; set; }

            public int LCID { get; set; }

            public string Location { get; set; }

            public string Name { get; set; }

            public string Description { get; set; }

            public override string ToString()
            {
                return string.Format(@"*\G{{{0}}}#{1}#{2}#{3}#{4}",
                    Guid,
                    Version,
                    LCID,
                    Location,
                    Description ?? Name);
            }

        }

        /// <summary>
        /// Describes a VB6 class item.
        /// </summary>
        class VB6ClassItem
        {

            /// <summary>
            /// Parses the given class definition.
            /// </summary>
            /// <param name="definition"></param>
            /// <returns></returns>
            public static VB6ClassItem Parse(string definition)
            {
                var a = definition.Split(new[] { "; " }, 2, StringSplitOptions.RemoveEmptyEntries);
                return new VB6ClassItem(a[0], a[1]);
            }

            /// <summary>
            /// Initializes a new instance.
            /// </summary>
            /// <param name="name"></param>
            /// <param name="file"></param>
            public VB6ClassItem(string name, string file)
            {
                Name = name;
                File = file;
            }

            /// <summary>
            /// Name of the class.
            /// </summary>
            public string Name { get; set; }

            /// <summary>
            /// Name of the file of the class.
            /// </summary>
            public string File { get; set; }

            /// <summary>
            /// Converts the object to a string.
            /// </summary>
            /// <returns></returns>
            public override string ToString()
            {
                return string.Format("{0}; {1}", Name, File);
            }

        }

        class VB6ModuleItem
        {

            /// <summary>
            /// Parses the given module definition.
            /// </summary>
            /// <param name="definition"></param>
            /// <returns></returns>
            public static VB6ModuleItem Parse(string definition)
            {
                var a = definition.Split(new[] { "; " }, 2, StringSplitOptions.RemoveEmptyEntries);
                return new VB6ModuleItem(a[0], a[1]);
            }

            /// <summary>
            /// Initializes a new instance.
            /// </summary>
            /// <param name="name"></param>
            /// <param name="file"></param>
            public VB6ModuleItem(string name, string file)
            {
                Name = name;
                File = file;
            }

            /// <summary>
            /// Name of the module.
            /// </summary>
            public string Name { get; set; }

            /// <summary>
            /// Name of the file of the module.
            /// </summary>
            public string File { get; set; }

            /// <summary>
            /// Converts the object to a string.
            /// </summary>
            /// <returns></returns>
            public override string ToString()
            {
                return string.Format("{0}; {1}", Name, File);
            }

        }

        class VB6FormItem
        {

            /// <summary>
            /// Parses the given form definition.
            /// </summary>
            /// <param name="definition"></param>
            /// <returns></returns>
            public static VB6FormItem Parse(string definition)
            {
                return new VB6FormItem(definition);
            }

            /// <summary>
            /// Initializes a new instance.
            /// </summary>
            /// <param name="file"></param>
            public VB6FormItem(string file)
            {
                File = file;
            }

            /// <summary>
            /// Name of the file of the form.
            /// </summary>
            public string File { get; set; }

            /// <summary>
            /// Converts the object to a string.
            /// </summary>
            /// <returns></returns>
            public override string ToString()
            {
                return string.Format("{0}", File);
            }

        }

        class VB6Project
        {

            /// <summary>
            /// Initializes a new instance.
            /// </summary>
            public VB6Project()
            {
                References = new List<VB6ReferenceItem>();
                Modules = new List<VB6ModuleItem>();
                Classes = new List<VB6ClassItem>();
                Forms = new List<VB6FormItem>();
                Properties = new Dictionary<string, object>();
                Sections = new Dictionary<string, Dictionary<string, object>>();
            }

            /// <summary>
            /// Gets or sets the type of the output.
            /// </summary>
            public VB6Type Type { get; set; }

            /// <summary>
            /// Project references.
            /// </summary>
            public List<VB6ReferenceItem> References { get; private set; }

            /// <summary>
            /// Included modules.
            /// </summary>
            public List<VB6ModuleItem> Modules { get; private set; }

            /// <summary>
            /// Included classes.
            /// </summary>
            public List<VB6ClassItem> Classes { get; private set; }

            /// <summary>
            /// Included forms.
            /// </summary>
            public List<VB6FormItem> Forms { get; private set; }

            /// <summary>
            /// Project properties.
            /// </summary>
            public Dictionary<string, object> Properties { get; private set; }

            /// <summary>
            /// Alternate section.
            /// </summary>
            public Dictionary<string, Dictionary<string, object>> Sections { get; private set; }

            /// <summary>
            /// Loads a <see cref="VB6Project"/> from the given file.
            /// </summary>
            /// <param name="path"></param>
            /// <returns></returns>
            public static VB6Project Load(string path)
            {
                var p = new VB6Project();

                foreach (var l in ParseVbp(path))
                {
                    // default project file options
                    if (string.IsNullOrWhiteSpace(l.Item1))
                    {
                        p.ApplyDefaultSectionLine(l.Item2, l.Item3);
                    }
                    else
                    {
                        // add section if not present
                        if (p.Sections.ContainsKey(l.Item1) == false)
                            p.Sections.Add(l.Item1, new Dictionary<string, object>());

                        // add property to section
                        p.Sections[l.Item1].Add(l.Item2, ParsePropertyValue(l.Item2, l.Item3));
                    }
                }

                return p;
            }

            /// <summary>
            /// Parses the VBP file by section and value.
            /// </summary>
            /// <returns></returns>
            public static IEnumerable<Tuple<string, string, string>> ParseVbp(string path)
            {
                var s = string.Empty;

                foreach (var l in File.ReadAllLines(path))
                {
                    var a = l.Trim();

                    // beginning of a new section
                    if (a.StartsWith("[") && a.EndsWith("]"))
                    {
                        s = a.Trim(new[] { '[', ']' });
                    }
                    else
                    {
                        var b = a.Split(new[] { '=' }, 2, StringSplitOptions.RemoveEmptyEntries);
                        if (b.Length >= 2)
                            yield return Tuple.Create(s, b[0].Trim(), b[1].Trim());
                    }
                }
            }

            /// <summary>
            /// Applies a key and value from the default section.
            /// </summary>
            /// <param name="key"></param>
            /// <param name="value"></param>
            void ApplyDefaultSectionLine(string key, string value)
            {
                if (string.IsNullOrWhiteSpace(key))
                    throw new ArgumentException("key");
                if (string.IsNullOrWhiteSpace(value))
                    throw new ArgumentException("value");

                switch (key)
                {
                    case "Type":
                        Type = (VB6Type)Enum.Parse(typeof(VB6Type), value);
                        break;
                    case "Reference":
                        References.Add(VB6ReferenceItem.Parse(value));
                        break;
                    case "Module":
                        Modules.Add(VB6ModuleItem.Parse(value));
                        break;
                    case "Class":
                        Classes.Add(VB6ClassItem.Parse(value));
                        break;
                    case "Form":
                        Forms.Add(VB6FormItem.Parse(value));
                        break;
                    default:
                        Properties.Add(key, ParsePropertyValue(key, value));
                        break;
                }
            }

            /// <summary>
            /// Parses the given property value.
            /// </summary>
            /// <param name="key"></param>
            /// <param name="value"></param>
            /// <returns></returns>
            static object ParsePropertyValue(string key, string value)
            {
                // quoted strings can be dequoted
                if (value.StartsWith("\"") &&
                    value.EndsWith("\""))
                    return value.Trim('"');

                // else try as an int
                int intValue;
                if (int.TryParse(value, out intValue))
                    return intValue;

                // else just trim and return
                return value.Trim();
            }

            /// <summary>
            /// Saves the project file to the given writer.
            /// </summary>
            /// <param name="path"></param>
            void Save(TextWriter writer)
            {
                writer.WriteLine("Type={0}", Type);

                foreach (var reference in References.OrderBy(i => i.Guid))
                    writer.WriteLine("Reference=" + reference.ToString());

                foreach (var module in Modules.OrderBy(i => i.Name))
                    writer.WriteLine("Module=" + module.ToString());

                foreach (var class_ in Classes.OrderBy(i => i.Name))
                    writer.WriteLine("Class=" + class_.ToString());

                foreach (var form in Forms.OrderBy(i => i.File))
                    writer.WriteLine("Form=" + form.ToString());

                foreach (var property in Properties.OrderBy(i => i.Key).ThenBy(i => i.Value))
                    writer.WriteLine(property.Key + "=" + PropertyValueToString(property.Key, property.Value));

                foreach (var section in Sections.OrderBy(i => i.Key))
                {
                    writer.WriteLine();
                    writer.WriteLine("[{0}]", section.Key);
                    foreach (var property in section.Value.OrderBy(i => i.Key))
                        writer.WriteLine(property.Key + "=" + PropertyValueToString(property.Key, property.Value));
                }
            }

            /// <summary>
            /// Saves the project file to the given path.
            /// </summary>
            /// <param name="path"></param>
            public void Save(string path)
            {
                // build project file output
                var e = new StringWriter();
                Save(e);
                var s = e.ToString();

                // write new file
                File.WriteAllText(path, s);
            }

            /// <summary>
            /// Converts a property value to a string.
            /// </summary>
            /// <param name="key"></param>
            /// <param name="value"></param>
            /// <returns></returns>
            string PropertyValueToString(string key, object value)
            {
                if (value is VB6Type)
                    return value.ToString();
                if (value is string)
                    return "\"" + value + "\"";
                if (value is int)
                    return value.ToString();

                throw new FormatException("Unhandled property value type.");
            }

        }

        /// <summary>
        /// Source VBP file to injest.
        /// </summary>
        [Required]
        public string Source { get; set; }

        /// <summary>
        /// Target VBP file to output.
        /// </summary>
        [Required]
        public string Target { get; set; }

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
        /// Applies the transforms.
        /// </summary>
        /// <param name="project"></param>
        /// <returns></returns>
        VB6Project TransformVbp(VB6Project project)
        {
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

            return project;
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
            int intValue;
            if (int.TryParse(value, out intValue))
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
            Log.LogMessage("Source: {0}", Source);
            Log.LogMessage("Target: {0}", Target);

            TransformVbp(VB6Project.Load(Source)).Save(Target);
            return true;
        }

    }



}
