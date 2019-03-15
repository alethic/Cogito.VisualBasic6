namespace Cogito.VisualBasic6.VB6C.Project
{

    public class VB6FormItem
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

}
