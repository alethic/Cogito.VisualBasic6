using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Cogito.Collections;

namespace Cogito.VisualBasic6.VB6C.Project
{

    /// <summary>
    /// Describes a VB6 profile.
    /// </summary>
    public class VB6Project
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
        /// Gets or sets the startup type of the application.
        /// </summary>
        public string Startup
        {
            get => (string)Properties.GetOrDefault("Startup");
            set => Properties["Startup"] = value?.TrimOrNull() ?? "(None)";
        }

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
            if (int.TryParse(value, out var intValue))
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

}
