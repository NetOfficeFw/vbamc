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


## When to use macros and why

There are several principal reasons to consider using Visual Basic
for Applications (VBA) macros in Microsoft Office.


### Automation and repetition

VBA is effective and efficient when it comes to repetitive solutions to formatting or correction problems. For example, have you ever changed the style of the paragraph at the top of each page in Word? Have you ever had to reformat multiple tables that were pasted from Excel into a Word document or an Outlook email? Have you ever had to make the same change in multiple Outlook contacts?

If you have a change that you have to make more than ten or twenty times, it may be worth automating it with VBA. If it is a change that you have to do hundreds of times, it certainly is worth considering. Almost any formatting or editing change that you can do by hand, can be done in VBA.

### Extensions to user interaction

There are times when you want to encourage or compel users to interact with the Office application or document in a particular way that is not part of the standard application. For example, you might want to prompt users to take some particular action when they open, save, or print a document.

### Interaction between Office applications

Do you need to copy all of your contacts from Outlook to Word and then format them in some particular way? Or, do you need to move data from Excel to a set of PowerPoint slides? Sometimes simple copy and paste does not do what you want it to do, or it is too slow. Use VBA programming to interact with the details of two or more Office applications at the same time and then modify the content in one application based on the content in another.

### Doing things another way

VBA programming is a powerful solution, but it is not always the optimal approach. Sometimes it makes sense to use other ways to achieve your aims.

The critical question to ask is whether there is an easier way. Before you begin a VBA project, consider the built-in tools and standard functionalities. For example, if you have a time-consuming editing or layout task, consider using styles or accelerator keys to solve the problem. Can you perform the task once and then use CTRL+Y (Redo) to repeat it? Can you create a new document with the correct format or template, and then copy the content into that new document?

Office applications are powerful; the solution that you need may already be there. Take some time to learn more about Office before you jump into programming.

Before you begin a VBA project, ensure that you have the time to work with VBA. Programming requires focus and can be unpredictable. Especially as a beginner, never turn to programming unless you have time to work carefully. Trying to write a "quick script" to solve a problem when a deadline looms can result in a very stressful situation. If you are in a rush, you might want to use conventional methods, even if they are monotonous and repetitive.


> Source: [Getting started with VBA in Office](https://learn.microsoft.com/en-us/office/vba/library-reference/concepts/getting-started-with-vba-in-office), under [CC-BY-4.0](https://github.com/MicrosoftDocs/VBA-Docs/blob/main/LICENSE) license.


## License

Source code is licensed under [MIT License](LICENSE.txt).

Copyright © 2022 Jozef Izso  
© 2022 Cisco Systems, Inc. All rights reserved.
