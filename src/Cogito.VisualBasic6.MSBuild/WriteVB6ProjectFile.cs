using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Cogito.VisualBasic6.VB6C.Project;

using Microsoft.Build.Framework;

namespace Cogito.VisualBasic6.MSBuild
{

    public class WriteVB6ProjectFile :
        Microsoft.Build.Utilities.Task
    {


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
            if (int.TryParse(value, out var intValue))
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
