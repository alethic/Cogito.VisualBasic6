using System;

namespace Cogito.VisualBasic6.VB6C.Project
{

    /// <summary>
    /// Describes a VB6 module.
    /// </summary>
    public class VB6ModuleItem
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

}
