using DocumentFormat.OpenXml;
using System;
using System.IO;
using System.Text;

namespace vbamc.Tests.Streams
{
    class VbaCompilerStreamsTests
    {
        [Test]
        public void Test1()
        {
            // Arrange
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var output = Path.Combine(TestContext.CurrentContext.TestDirectory, DateTimeOffset.Now.ToUnixTimeSeconds().ToString());
            Directory.CreateDirectory(output);

            var sourcePath = Path.Combine(TestContext.CurrentContext.TestDirectory, "data");
            var classPath = Path.Combine(sourcePath, "Class.vb");
            var modulePath = Path.Combine(sourcePath, "Module.vb");

            var compiler = new VbaCompiler()
            {
                ProjectId = Guid.NewGuid(),
                ProjectName = "Project A",
                ProjectVersion = "1.0.0",
                CompanyName = "ACME"
            };
            compiler.AddModule(modulePath);
            compiler.AddClass(classPath);

            var vbaProjectMemory = new MemoryStream();
            var powerpointAddinMemory = new MemoryStream();

            // Act
            compiler.CompileVbaProject(vbaProjectMemory);
            compiler.CompilePowerPointMacroFile(powerpointAddinMemory, vbaProjectMemory, PresentationDocumentType.MacroEnabledPresentation);

            // Assert
            ClassicAssert.Greater(powerpointAddinMemory.Length, 1);
        }
    }
}
