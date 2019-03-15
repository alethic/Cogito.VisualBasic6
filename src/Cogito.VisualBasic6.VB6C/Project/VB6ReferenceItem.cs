using System;
using System.Text.RegularExpressions;

namespace Cogito.VisualBasic6.VB6C.Project
{

    /// <summary>
    /// Describes a referenec item within a <see cref="VB6Project"/>.
    /// </summary>
    public class VB6ReferenceItem
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

}
