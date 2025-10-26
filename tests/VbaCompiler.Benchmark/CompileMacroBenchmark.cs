using System.Text;
using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml;
using vbamc;

public class CompileMacroBenchmark
{
    private readonly Guid ProjectId = Guid.NewGuid();

    [GlobalSetup]
    public void Setup()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    [Benchmark]
    public MemoryStream CompilePowerPointAddinMacro()
    {
        // Arrange
        var sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "data");
        var classPath = Path.Combine(sourcePath, "Class.vb");
        var modulePath = Path.Combine(sourcePath, "Module.vb");

        var compiler = new VbaCompiler()
        {
            ProjectId = this.ProjectId,
            ProjectName = "Project A",
            ProjectVersion = "1.0.0",
            CompanyName = "ACME"
        };
        compiler.AddModule(modulePath);
        compiler.AddClass(classPath);

        // Act
        var macro = new MemoryStream();
        var vbaProjectStream = compiler.CompileVbaProject();
        compiler.CompilePowerPointMacroFile(macro, vbaProjectStream, PresentationDocumentType.AddIn);
        return macro;
    }
}
