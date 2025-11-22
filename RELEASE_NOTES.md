# v2.0.0

VbaCompiler v2.0.0 is a major release that brings support for .NET 10, modernizes dependencies, and improves performance.

### What's New
- The `VbaCompiler` library and `vbamc` tool work on .NET 8 and 10 runtimes.
- Works on Linux, macOS, and Windows.


### Installation

Install as a .NET tool:

```bash
dotnet tool install --global vbamc --version 2.0.0
```

Upgrade from previous version:

```bash
dotnet tool update --global vbamc
```

Add `VbaCompiler` library to your project:

```bash
dotnet add package NetOfficeFw.VbaCompiler --version 2.0.0
```


### Usage

Use `vbamc` tool to generate a PowerPoint macro-enabled presentation:

```bash
vbamc --module Module.vb --class MyClass.vb --name "VBA Project" --company "ACME" --file Presentation.ppam
```

#### Using VbaCompiler library

```csharp
using DocumentFormat.OpenXml.Packaging;
using vbamc;

var compiler = new VbaCompiler();
compiler.ProjectId = Guid.NewGuid();
compiler.ProjectName = "My Macro Project";
compiler.CompanyName = "ACME";
compiler.AddModule("Module.vb");
compiler.AddClass("MyClass.vb");

// Generate the vbaProject.bin file
var vbaProjectPath = compiler.CompileVbaProject("obj", "vbaProject.bin");

// Generate Excel macro file
var macroFilePath = compiler.CompileExcelMacroFile(
    "bin", 
    "MyMacro.xlsm", 
    vbaProjectPath, 
    SpreadsheetDocumentType.MacroEnabledWorkbook
);
```


### Dependency Updates
- **OpenMcdf 3.1.0**: Updated to the latest version of OpenMcdf library for improved CFB (Compound File Binary) file format handling.
- **DocumentFormat.OpenXml 3.3.0**: Latest version provides better Office Open XML document processing.
- **NetOfficeFw.VbaCompression 3.0.1**: Updated compression library for VBA project files.
- **Microsoft.SourceLink.GitHub 8.0.0**: Enhanced debugging experience with up-to-date source linking.

### Infrastructure Enhancements
- **Modernized CI/CD**: Updated release workflow with Trusted Publishing for more secure package distribution.
- **Artifact attestation**: Enhanced security with artifact attestation in the release process.
- **Automated dependency updates**: Configured Dependabot for GitHub Actions to keep CI/CD dependencies current.

### Breaking Changes

This major release removes support for .NET 6 and .NET 7. The project now targets only .NET 8 and .NET 10.

### Acknowledgments

Thank you to all contributors who helped make this release possible!
