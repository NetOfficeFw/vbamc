using System.Text;
using vbamc;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

var targetPath = @"vbaProject.bin";

var compiler = new VbaCompiler();

compiler.ProjectId = new Guid("{607B0672-60F3-4698-B08A-FA8F558E7F13}");
compiler.ProjectName = "CustomProjectName";

compiler.AddModule(@"d:\dev\github\NetOfficeFw\vbamc\sample\Module1.vb");
compiler.AddClass(@"d:\dev\github\NetOfficeFw\vbamc\sample\Class1.vb");

compiler.Compile(targetPath);
