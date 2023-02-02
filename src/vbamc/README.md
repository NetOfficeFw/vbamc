Compiler `vbamc` compiles Visual Basic source code and Ribbon customizations to a macro enabled add-in file for Microsoft Office applications.

### Usage

Sample usage to generate PowerPoint presentation file with macro
from the source code files `Module.vb` and `MyClass.vb`:

```shell
vbamc --module Module.vb --class MyClass.vb --name "VBA Project" --company "ACME" --file Presentation.ppam
```

### Requirements

The compiler works on .NET 6 and .NET 7 runtimes on Windows and macOS.

### Samples

Discover samples in our repository at <https://github.com/NetOfficeFw/vbamc/tree/main/sample>


_Project icon [Online Coding][1] is licensed from Icons8 service under Universal Multimedia Licensing Agreement for Icons8._  
_See <https://icons8.com/license> for more information_

[1]: https://icons8.com/icon/UVQTFk728g0D/online-coding
