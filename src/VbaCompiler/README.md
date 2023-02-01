Library `VbaCompiler` compiles Visual Basic source code and Ribbon customizations to Microsoft Office macro enabled files.

### Usage

Sample usage to generate Excel workbook file with macro
from the source code files `Module.vb` and `MyClass.vb`:

```csharp
using DocumentFormat.OpenXml.Packaging;

var compiler = new VbaCompiler();
compiler.ProjectId = Guid.NewGuid();
compiler.ProjectName = "My Macro Project";
compiler.CompanyName = "ACME";
compiler.AddModule("Module.vb");
compiler.AddClass("MyClass.vb");

// generate the vbaProject.bin file
var vbaProjectPath = compiler.CompileVbaProject("obj", "vbaProject.bin");

// generate Excel macro file
var macroFile = compiler.CompileExcelMacroFile("bin", "MyMacro.xlsm", vbaProjectPath, SpreadsheetDocumentType.MacroEnabledWorkbook);
```

### Requirements

The compiler works on .NET 6 and .NET 7 runtimes on Windows and macOS.

### Samples

Discover samples in our repository at <https://github.com/NetOfficeFw/vbamc/tree/main/sample>


_Project icon [Code][1] is licensed from Icons8 service under Universal Multimedia Licensing Agreement for Icons8._  
_See <https://icons8.com/license> for more information_

[1]: https://icons8.com/icon/43988/code
