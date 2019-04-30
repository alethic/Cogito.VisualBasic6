using System.Diagnostics;
using System.IO;

using Microsoft.Build.Framework;

namespace Cogito.VisualBasic6.MSBuild
{

    public class GetFileVersionInfo :
        Microsoft.Build.Utilities.Task
    {

        [Output]
        public ITaskItem[] Files { get; set; }

        public override bool Execute()
        {
            if (Files != null)
            {
                foreach (var f in Files)
                {
                    if (File.Exists(f.ItemSpec))
                    {
                        var v = FileVersionInfo.GetVersionInfo(f.ItemSpec);
                        f.SetMetadata(nameof(FileVersionInfo.Comments), v.Comments);
                        f.SetMetadata(nameof(FileVersionInfo.CompanyName), v.CompanyName);
                        f.SetMetadata(nameof(FileVersionInfo.FileBuildPart), v.FileBuildPart.ToString());
                        f.SetMetadata(nameof(FileVersionInfo.FileDescription), v.FileDescription);
                        f.SetMetadata(nameof(FileVersionInfo.FileMajorPart), v.FileMajorPart.ToString());
                        f.SetMetadata(nameof(FileVersionInfo.FileMinorPart), v.FileMinorPart.ToString());
                        f.SetMetadata(nameof(FileVersionInfo.FilePrivatePart), v.FilePrivatePart.ToString());
                        f.SetMetadata(nameof(FileVersionInfo.FileVersion), v.FileVersion);
                        f.SetMetadata(nameof(FileVersionInfo.InternalName), v.InternalName);
                        f.SetMetadata(nameof(FileVersionInfo.IsDebug), v.IsDebug.ToString());
                        f.SetMetadata(nameof(FileVersionInfo.IsPatched), v.IsPatched.ToString());
                        f.SetMetadata(nameof(FileVersionInfo.IsPreRelease), v.IsPreRelease.ToString());
                        f.SetMetadata(nameof(FileVersionInfo.IsPrivateBuild), v.IsPrivateBuild.ToString());
                        f.SetMetadata(nameof(FileVersionInfo.IsSpecialBuild), v.IsSpecialBuild.ToString());
                        f.SetMetadata(nameof(FileVersionInfo.Language), v.Language);
                        f.SetMetadata(nameof(FileVersionInfo.LegalCopyright), v.LegalCopyright);
                        f.SetMetadata(nameof(FileVersionInfo.LegalTrademarks), v.LegalTrademarks);
                        f.SetMetadata(nameof(FileVersionInfo.OriginalFilename), v.OriginalFilename);
                        f.SetMetadata(nameof(FileVersionInfo.PrivateBuild), v.PrivateBuild);
                        f.SetMetadata(nameof(FileVersionInfo.ProductBuildPart), v.ProductBuildPart.ToString());
                        f.SetMetadata(nameof(FileVersionInfo.ProductMajorPart), v.ProductMajorPart.ToString());
                        f.SetMetadata(nameof(FileVersionInfo.ProductName), v.ProductName);
                        f.SetMetadata(nameof(FileVersionInfo.ProductPrivatePart), v.ProductPrivatePart.ToString());
                        f.SetMetadata(nameof(FileVersionInfo.ProductVersion), v.ProductVersion);
                        f.SetMetadata(nameof(FileVersionInfo.SpecialBuild), v.SpecialBuild);
                    }
                }
            }

            return true;
        }

    }

}
