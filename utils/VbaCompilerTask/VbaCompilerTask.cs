using DocumentFormat.OpenXml.Packaging;
using System.Text;
using DocumentFormat.OpenXml.Office2010.CustomUI;
using Microsoft.Build.Framework;
using vbamc;
using Task = Microsoft.Build.Utilities.Task;

// ReSharper disable once UnusedMember.Global
public class VbaCompilerTask : Task
{
    [Required]
    public ITaskItem[] Modules { get; set; } = null!;

    [Required]
    public ITaskItem[] Classes { get; set; } = null!;

    [Required]
    public string ProjectName { get; set; } = null!;

    [Required]
    public string OutputName { get; set; } = null!;

    [Required]
    public string CompanyName { get; set; } = null!;

    [Output]
    public string OutputMacroFile { get; set; } = null!;

    public override bool Execute()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        var wd = Directory.GetCurrentDirectory();
        var outputPath = Path.Combine(wd, "bin");
        var intermediatePath = Path.Combine(wd, "obj");

        var outputProjectName = @"vbaProject.bin";
        var outputFileName = this.OutputName;

        DirectoryEx.EnsureDirectory(outputPath);
        var targetMacroPath = Path.Combine(outputPath, outputFileName);

        var compiler = new VbaCompiler();

        compiler.ProjectId = Guid.NewGuid();
        compiler.ProjectName = this.ProjectName;

        // add modules
        foreach (var module in this.Modules)
        {
            var path = module.ItemSpec;
            if (!Path.IsPathRooted(path))
            {
                path = Path.Combine(wd, path);
            }

            compiler.AddModule(path);
        }

        // add classes
        foreach (var @class in this.Classes)
        {
            var path = @class.ItemSpec;
            if (!Path.IsPathRooted(path))
            {
                path = Path.Combine(wd, path);
            }

            compiler.AddClass(path);
        }

        var projectPath = compiler.Compile(intermediatePath, outputProjectName);

        var baseDirectory = Path.GetDirectoryName(this.GetType().Assembly.Location);

        var macroTemplatePath = Path.Combine(baseDirectory!, @"data/MacroTemplate.potm");
        var macroTemplate = PresentationDocument.CreateFromTemplate(macroTemplatePath);
        var mainDoc = macroTemplate.PresentationPart;
        if (mainDoc != null)
        {
            var vbaProject = mainDoc.AddNewPart<VbaProjectPart>();
            using var reader = File.OpenRead(projectPath);
            vbaProject.FeedData(reader);
            reader.Close();
        }

        AttachRibbonCustomization(macroTemplate, Directory.GetCurrentDirectory());

        macroTemplate.PackageProperties.Title = this.ProjectName;
        var propCompany = macroTemplate.ExtendedFilePropertiesPart?.Properties.Company;
        if (propCompany != null && !String.IsNullOrEmpty(this.CompanyName))
        {
            propCompany.Text = this.CompanyName;
        }

        macroTemplate.ChangeDocumentType(DocumentFormat.OpenXml.PresentationDocumentType.AddIn);
        macroTemplate.SaveAs(targetMacroPath);

        this.OutputMacroFile = targetMacroPath;

        return true;
    }

    private void AttachRibbonCustomization(PresentationDocument document, string sourcePath)
    {
        var customUiDir = Path.Combine(sourcePath, "customUI");
        var ribbonPath = Path.Combine(customUiDir, "customUI14.xml");
        if (!File.Exists(ribbonPath))
        {
            return;
        }

        var ribbonContent = File.ReadAllText(ribbonPath);

        var ribbonPart = document.RibbonAndBackstageCustomizationsPart;
        if (ribbonPart == null)
        {
            ribbonPart = document.AddRibbonAndBackstageCustomizationsPart();
        }

        ribbonPart.CustomUI = new CustomUI(ribbonContent);
        ribbonPart.CustomUI.Save();
        Console.WriteLine($"Added ribbon customization from file '{ribbonPath}'");

        var images = Directory.EnumerateFiles(Path.Combine(customUiDir, "images"), "*.png");
        foreach (var imagePath in images)
        {
            var imageFilename = Path.GetFileNameWithoutExtension(imagePath);
            var imagePart = ribbonPart.AddImagePart(ImagePartType.Png, imageFilename);
            using var imageStream = new FileStream(imagePath, FileMode.Open);
            imagePart.FeedData(imageStream);
        }
    }
}
