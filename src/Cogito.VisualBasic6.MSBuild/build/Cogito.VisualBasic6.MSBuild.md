# Reg-Free COM Build Extensions

The `COM.targets` and `COM.props` extensions are used to augment the compilation of standard SDK-based C# projects that generate COM outputs.

Two properties are available for projects: `EnableCOMExport` and `EnableCOMImport`. These properties and build extensions work together to enable registration-free COM situations.

`EnableCOMExport` augments the build process to generate and embed TypeLibs and registration-free COM assembly manifests within the generated C# assemblies. Set this property on projects which generate COM libraries. The embedded manifests and type libraries can then be used with COM applications for discovery.

`EnableCOMImport` augments the build process for executables to generate and embed registration-free COM application manifests within the output executable. This causes the COM assembly context generation to search for additional manifests when resolving COM objects.