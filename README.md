# Cogito.VisualBasic6.MSBuild

MSBuild development time support for building Visual Basic 6 projects. This package introduces additional MSBuild targets and properties for handling Visual Basic 6 project builds.

## Dependencies

This package requires VB6 to be installed on your machine. The path to VB6 is discovered by using the `HKLM\SOFTWARE\Microsoft\VisualStudio\6.0` registry keys. The MSBuild property `VB6ProductDir` can be set if this detection logic is not sufficient.

The package also requires a Windows SDK containing the `mt.exe` utility, and the Microsft Visual C++ build components containing the appropriate build tasks. These are usually found by installing C++ related portions of Visual Studio 2017.

## Usage

Install this package within a standard 2017 .csproj project. Create a .vbp file as content within your .NET project. Make sure it is NOT named the same as your .NET project. Usually VB6 project names do not contain periods or namespaces. Two outputs are generated from the build: the original .Net DLL, and a new VB6 DLL or EXE.

Some MSBuild properties are required:

+ `VB6ProjectFile`: name of the .vbp file. By default this will be the same as your C# project name with a `.vbp` extension.
+ `VB6Name`: Name of the VB6 project. This ends up being the ProgID for ActiveX DLLs.
+ `VB6ExeName`: Output `.dll` name. Defaults to  `VB6Name` with .dll extension.

VB6 resources to be built include Modules, Classes and Forms. Three `ItemGroup`s are used for this purpose:

+ `VB6Module`
+ `VB6Class`
+ `VB6Form`

Files with `.bas`, `.cls` and `.frm` extensions are automatically added to these `ItemGroup`s.

All `PackageReference` or `ProjectReference` or `COMReference` items within your `.csproj` are scanned to build the set of references passed to VB6. That means you can add a project reference to another VB6 project, and it's output will be included as a reference. You can use the standard Visual Studio COM references for traditional COM components. Any dependent .NET components with COM manifests are also available as references.

The original `.vbp` file is not actually used for build. Instead, it is parsed and augmented with the metadata from the C# project. VB6.exe is then used to silently execute the generated temporary file.

The generated VB6 DLL files contain embedded COM assembly manifests in the `RT_MANIFEST;2` resource. This allows for registration-free COM usage. It is also important for building references between projects as these manifests are used to detect depedencies.

Additionally, upon successful build, a COM Interop assembly for the VB6 DLL is generated. This interop assembly is then merged into the main output assembly of your C# project.

Consult the `Cogito.COM.MSBuild` package for more information on the COM manifest infrastructure.

## Internals

The NuGet package includes a MSBuild task `VB6C` for compiling VB project data. This invokes a `Compiler` class. The `Compiler` class generates a temporary `.vbp` file, and then spawns a `Cogito.VisualBasic6.VB6C` executable, feeding it paths to the temporary files. This executable uses `EasyHook` to spawn the `VB6.exe` executable, injecting some overrides into it. It silences the annoying VB6 beep sounds. And it WILL but not yet, configure the COM ActCtx. It then relays the exit code back to the main process.

The VB6C executable collections the VB6 error output, and emits it to Standard Error. As a compiler executable should.

The Compiler class traps this, and emits it as MSBuild error log messages. And fails the compile.
