// Copyright 2022 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

using System;
using System.ComponentModel.DataAnnotations;
using System.Reflection;
using System.Text;
using DocumentFormat.OpenXml.Office2010.CustomUI;
using DocumentFormat.OpenXml.Packaging;
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

    [Option("-t|--type", Description = "Project name")]
    public IEnumerable<string> TargetTypes { get; } = new [] { "Excel", "PowerPoint" };

    [Option("-n|--name", Description = "Project name")]
    public string ProjectName { get; } = "VBAProject";

    [Option("--company", Description = "Company name")]
    public string? CompanyName { get; } = null!;

    [Option("-f|--file", Description = "Target add-in file name without extension")]
    public string FileName { get; } = "Addin";

    [Option("-o|--output", Description = "Target build output path")]
    public string OutputPath { get; } = "bin";

    [Option("--intermediate", Description = "Intermediate path for build output")]
    public string IntermediatePath { get; } = "obj";

    private void OnExecute()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        var wd = Directory.GetCurrentDirectory();
        var outputPath = Path.Combine(wd, this.OutputPath);
        var intermediatePath = Path.Combine(wd, this.IntermediatePath);

        var outputProjectName = @"vbaProject.bin";

        DirectoryEx.EnsureDirectory(outputPath);

        var compiler = new VbaCompiler();

        compiler.ProjectId = Guid.NewGuid();
        compiler.ProjectName = this.ProjectName;

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

        var projectPath = compiler.Compile(intermediatePath, outputProjectName);

        foreach (var targetType in this.TargetTypes)
        {
            switch (targetType)
            {
                case "PowerPoint":
                    GeneratePowerPointAddinFile(outputPath, projectPath);
                    break;
                case "Excel":
                    GenerateExcelAddinFile(outputPath, projectPath);
                    break;
            }
        }
    }

    private string GeneratePowerPointAddinFile(string outputPath, string vbaProjectFilePath)
    {
        var outputFileName = Path.ChangeExtension(this.FileName, ".ppam");
        var targetMacroPath = Path.Combine(outputPath, outputFileName);

        var macroTemplatePath = Path.Combine(AppContext.BaseDirectory, @"data/MacroTemplate.potm");
        var macroTemplate = PresentationDocument.CreateFromTemplate(macroTemplatePath);
        var mainDoc = macroTemplate.PresentationPart;
        if (mainDoc != null)
        {
            var vbaProject = mainDoc.AddNewPart<VbaProjectPart>();
            using var reader = File.OpenRead(vbaProjectFilePath);
            vbaProject.FeedData(reader);
            reader.Close();
        }

        var ribbonPart = macroTemplate.RibbonAndBackstageCustomizationsPart ?? macroTemplate.AddRibbonAndBackstageCustomizationsPart();

        AttachRibbonCustomization(ribbonPart, Directory.GetCurrentDirectory());

        macroTemplate.PackageProperties.Title = this.ProjectName;
        var propCompany = macroTemplate.ExtendedFilePropertiesPart?.Properties.Company;
        if (propCompany != null && !String.IsNullOrEmpty(this.CompanyName))
        {
            propCompany.Text = this.CompanyName;
        }

        macroTemplate.ChangeDocumentType(DocumentFormat.OpenXml.PresentationDocumentType.AddIn);
        macroTemplate.SaveAs(targetMacroPath);

        Console.WriteLine($"Generated {outputFileName} add-in file.");

        return targetMacroPath;
    }

    private string GenerateExcelAddinFile(string outputPath, string vbaProjectFilePath)
    {
        var outputFileName = Path.ChangeExtension(this.FileName, ".xlam");
        var targetMacroPath = Path.Combine(outputPath, outputFileName);

        var macroTemplatePath = Path.Combine(AppContext.BaseDirectory, @"data/MacroTemplate.xltm");
        var macroTemplate = SpreadsheetDocument.CreateFromTemplate(macroTemplatePath);
        TypedOpenXmlPart? mainDoc = macroTemplate.WorkbookPart;
        if (mainDoc != null)
        {
            var vbaProject = mainDoc.AddNewPart<VbaProjectPart>();
            using var reader = File.OpenRead(vbaProjectFilePath);
            vbaProject.FeedData(reader);
            reader.Close();
        }

        var ribbonPart = macroTemplate.RibbonAndBackstageCustomizationsPart ?? macroTemplate.AddRibbonAndBackstageCustomizationsPart();
        AttachRibbonCustomization(ribbonPart, Directory.GetCurrentDirectory());

        macroTemplate.PackageProperties.Title = this.ProjectName;
        var propCompany = macroTemplate.ExtendedFilePropertiesPart?.Properties.Company;
        if (propCompany != null && !String.IsNullOrEmpty(this.CompanyName))
        {
            propCompany.Text = this.CompanyName;
        }

        macroTemplate.ChangeDocumentType(DocumentFormat.OpenXml.SpreadsheetDocumentType.AddIn);
        macroTemplate.SaveAs(targetMacroPath);

        Console.WriteLine($"Generated {outputFileName} add-in file.");

        return targetMacroPath;
    }


    private void AttachRibbonCustomization(RibbonAndBackstageCustomizationsPart ribbonPart, string sourcePath)
    {
        var customUiDir = Path.Combine(sourcePath, "customUI");
        var ribbonPath = Path.Combine(customUiDir, "customUI14.xml");
        if (!File.Exists(ribbonPath))
        {
            return;
        }

        var ribbonContent = File.ReadAllText(ribbonPath);

        ribbonPart.CustomUI = new CustomUI(ribbonContent);
        ribbonPart.CustomUI.Save();

        var images = Directory.EnumerateFiles(Path.Combine(customUiDir, "images"), "*.png");
        foreach(var imagePath in images)
        {
            var imageFilename = Path.GetFileNameWithoutExtension(imagePath);
            var imagePart = ribbonPart.AddImagePart(ImagePartType.Png, imageFilename);
            using var imageStream = new FileStream(imagePath, FileMode.Open);
            imagePart.FeedData(imageStream);
        }
    }
}
