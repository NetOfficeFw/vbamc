# vbamc

> Compile macro enabled add-ins for Microsoft Office applications.

Compiler `vbamc` compiles Visual Basic source code and Ribbon customizations to Microsoft Office macro files or macro enabled add-ins. It supports Microsoft Word, Excel and PowerPoint.


## Installation

Use `dotnet tool` command to install the `vbamc` compiler:

```commandline
dotnet tool install --global vbamc
```


## Usage

Pass the list of source code files with modules and classes to the compiler:

```commandline
vbamc -m Module.vb -c MyClass.vb -f AcmeSample -n "Sample Addin" --company "ACME"
```

It will generate the macro files named `AcmeSampleMacro.{docm,xlsm,pptm}` usable
in Microsoft Word, Excel and PowerPoint.


## License

Source code is licensed under [MIT License](LICENSE.txt).

Copyright © 2022 Jozef Izso  
© 2022 Cisco Systems, Inc. All rights reserved.
