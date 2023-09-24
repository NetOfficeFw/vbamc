// Copyright 2022 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

using System.ComponentModel.DataAnnotations;
using System.Text;
using DocumentFormat.OpenXml;
using McMaster.Extensions.CommandLineUtils;
using vbamc;

[Command(Name = "vbamc", Description = "Visual Basic for Applications macro compiler")]
public class Program
{
    public static int Main(string[] args)
        => CommandLineApplication.Execute<Program>(args);

    [Required]
    [Option("-m|--module")]
    public IEnumerable<string> Modules { get; } = Enumerable.Empty<string>();

    [Option("-c|--class")]
    public IEnumerable<string> Classes { get; } = Enumerable.Empty<string>();

    // [Option("-d|--document")]
    // public string? Document { get; }

    [Option("-n|--name", Description = "Project name")]
    public string ProjectName { get; } = "VBAProject";

    [Option("--company", Description = "Company name")]
    public string? CompanyName { get; }

    [Option("-f|--file", Description = "Target add-in file name with extension")]
    public string FileName { get; } = "PresentationAddin.ppam";

    [Option("-o|--output", Description = "Target build output path")]
    public string OutputPath { get; } = "bin";

    [Option("--user-profile-path", Description = "Path to the user profile to replace the ~/ expression")]
    public string? UserProfilePath { get; }

    private void OnExecute()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        var wd = Directory.GetCurrentDirectory();
        var outputPath = Path.Combine(wd, this.OutputPath);

        var compiler = new VbaCompiler();

        compiler.ProjectId = Guid.NewGuid();
        compiler.ProjectName = this.ProjectName;
        compiler.CompanyName = this.CompanyName;
        compiler.UserProfilePath = this.UserProfilePath;

        // // add document module
        // if (this.Document != null)
        // {
        //     var path = this.Document;
        //     if (!Path.IsPathRooted(path))
        //     {
        //         path = Path.Combine(wd, path);
        //     }

        //     compiler.AddThisDocument(path);
        // }

        // add modules
        foreach (var module in this.Modules)
        {
            var path = module;
            if (!Path.IsPathRooted(path))
            {
                path = Path.Combine(wd, path);
            }

            compiler.AddModule(path);
        }

        // add classes
        foreach (var @class in this.Classes)
        {
            var path = @class;
            if (!Path.IsPathRooted(path))
            {
                path = Path.Combine(wd, path);
            }

            compiler.AddClass(path);
        }


        DirectoryEx.EnsureDirectory(outputPath);
        using var outputMacroFile = File.Create(Path.Combine(outputPath, this.FileName));
        var vbaProjectMemory = compiler.CompileVbaProject();

        var extension = Path.GetExtension(this.FileName).ToLowerInvariant();
        switch (extension)
        {
            // Microsoft PowerPoint
            case ".pptm":
                compiler.CompilePowerPointMacroFile(outputMacroFile, vbaProjectMemory, PresentationDocumentType.MacroEnabledPresentation);
                break;
            case ".ppam":
                compiler.CompilePowerPointMacroFile(outputMacroFile, vbaProjectMemory, PresentationDocumentType.AddIn);
                break;

            // Microsoft Excel
            case ".xlsm":
                compiler.CompileExcelMacroFile(outputMacroFile, vbaProjectMemory, SpreadsheetDocumentType.MacroEnabledWorkbook);
                break;
            case ".xlam":
                compiler.CompileExcelMacroFile(outputMacroFile, vbaProjectMemory, SpreadsheetDocumentType.AddIn);
                break;

            // Microsoft Word
            case ".docm":
                compiler.CompileWordMacroFile(outputMacroFile, vbaProjectMemory, WordprocessingDocumentType.MacroEnabledDocument);
                break;

            default:
                throw new NotSupportedException($"File extension {extension} is not supported.");
        }
    }
}
