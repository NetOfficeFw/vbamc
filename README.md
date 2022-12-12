# vbamc

> Compile macro enabled add-ins for Microsoft Office applications.

Compiler `vbamc` compiles Visual Basic source code and Ribbon customizations to a macro
enabled add-in file for Microsoft Office applications.


## Installation

Use `dotnet tool` command to install the `vbamc` compiler:

```commandline
dotnet tool install --global vbamc
```


## Usage

Pass the list of source code files with modules and classes to the compiler:

```commandline
vbamc -m Module.vb -c MyClass.vb -n "Sample Addin" --company "ACME"
```

It will generate the macro file.


## License

Source code is licensed under [MIT License](LICENSE.txt).

Copyright © 2022 Jozef Izso  
© 2022 Cisco Systems, Inc. All rights reserved.
