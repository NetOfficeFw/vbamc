using DocumentFormat.OpenXml.Packaging;
using System.Text;
using vbamc;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

var targetPath = @"vbaProject.bin";
var targetMacroPath = @"CustomMacro.pptm";

var compiler = new VbaCompiler();

compiler.ProjectId = Guid.NewGuid();
compiler.ProjectName = "CustomProjectName";

compiler.AddModule(@"d:\dev\github\NetOfficeFw\vbamc\sample\Module1.vb");
compiler.AddClass(@"d:\dev\github\NetOfficeFw\vbamc\sample\Class1.vb");

compiler.Compile(targetPath);

var macroTemplatePath = Path.Combine(Directory.GetCurrentDirectory(), @"data\MacroTemplate.potm");
var macroTemplate = PresentationDocument.CreateFromTemplate(macroTemplatePath);
var mainDoc = macroTemplate.PresentationPart;
if (mainDoc != null)
{
    var vbaProject = mainDoc.AddNewPart<VbaProjectPart>();
    using var reader = File.OpenRead(targetPath);
    vbaProject.FeedData(reader);
    reader.Close();
}

macroTemplate.ChangeDocumentType(DocumentFormat.OpenXml.PresentationDocumentType.MacroEnabledPresentation);
macroTemplate.SaveAs(targetMacroPath);