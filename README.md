# Cogito.VisualBasic6.MSBuild

MSBuild development time support for building Visual Basic 6 projects.

Install this package within a standard 2017 .csproj project to insert build steps for generating VB6 ActiveX libraries.
Include a .vbp file as an item in your project. There are a number of MSBuild properties you should set:

VB6ProjectFile: name of the .vbp file. By default this will be the same as your C# project name, with a .vbp extension
VB6Name: Name of the ActiveX project.
VB6ExeName: Output .DLL name. Defaults to the VB6Name with .dll extension.

ItemGroups for VB6Class and VB6Module are included by default.

During build your VBP file is used as a source for a dynamically generated VBP file. This new VBP file has all of your
native VS references configured within it. That includes ProjectReferences, PackageReferences, and COMReferences. It also
considers TypeLibs within any Content files of dependent projects.

Additionally, the generated TLB is merged into the output DLL as Win32 resource RT_MANIFEST, for Reg-Free discovery.

You can have one VB6 project depend on another. Each one scans the VS references to the others and configures the hint paths.

Generated output files are configured for Reg-Free COM activation.

Consule the Cogito.COM.MSBuild package for more information on that.
