using System;

namespace Cogito.VisualBasic6.VB6C.Project
{

    /// <summary>
    /// Describes a VB6 class item.
    /// </summary>
    public class VB6ClassItem
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

}
