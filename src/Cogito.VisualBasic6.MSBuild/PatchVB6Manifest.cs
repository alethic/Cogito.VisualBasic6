using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;

using Microsoft.Build.Framework;

namespace Cogito.VisualBasic6.MSBuild
{

    /// <summary>
    /// Applies various patches to the VB6 assembly manifest.
    /// </summary>
    public class PatchVB6Manifest :
        Microsoft.Build.Utilities.Task
    {

        [Required]
        public string ManifestFile { get; set; }

        [Required]
        public string VBRegFile { get; set; }

        public override bool Execute()
        {
            var asm = (XNamespace)"urn:schemas-microsoft-com:asm.v1";

            // load existing manifest file
            var xml = XDocument.Load(ManifestFile);

            // parse VB registration file
            var reg = File.ReadAllLines(VBRegFile)
                .Select(i => i.Split(new[] { " = " }, StringSplitOptions.RemoveEmptyEntries))
                .Where(i => i.Length == 2)
                .ToDictionary(i => i[0], i => i[1]);

            foreach (var comClass in xml.Descendants(asm + "comClass"))
            {
                var clsid = (string)comClass.Attribute("clsid");
                if (clsid == null)
                    continue;

                var key = @"HKEY_CLASSES_ROOT\CLSID\" + clsid + @"\ProgID";
                if (reg.TryGetValue(key, out var progid))
                    comClass.SetAttributeValue("progid", progid);

                comClass.SetAttributeValue("threadingModel", "Apartment");
            }

            xml.Save(ManifestFile);

            return true;
        }

    }

}
