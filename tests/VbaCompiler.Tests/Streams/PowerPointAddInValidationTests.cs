using System;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace vbamc.Tests.Streams
{
    /// <summary>
    /// Integration tests that validate generated PowerPoint add-in packages
    /// have correct OpenXML structure and contain valid VBA project parts.
    /// Addresses issue #40: https://github.com/NetOfficeFw/vbamc/issues/40
    /// </summary>
    class PowerPointAddInValidationTests
    {
        [SetUp]
        public void Setup()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        private static (string modulePath, string classPath) GetTestDataPaths()
        {
            var sourcePath = Path.Combine(TestContext.CurrentContext.TestDirectory, "data");
            var classPath = Path.Combine(sourcePath, "Class.vb");
            var modulePath = Path.Combine(sourcePath, "Module.vb");
            return (modulePath, classPath);
        }

        private static VbaCompiler CreateTestCompiler(string projectName)
        {
            return new VbaCompiler
            {
                ProjectId = Guid.NewGuid(),
                ProjectName = projectName,
                ProjectVersion = "1.0.0",
                CompanyName = "TestCompany"
            };
        }

        private static MemoryStream CompilePowerPointFile(
            VbaCompiler compiler,
            PresentationDocumentType documentType)
        {
            var outputStream = new MemoryStream();

            // Use the parameterless overload that returns a stream with position already at 0
            var vbaProjectStream = compiler.CompileVbaProject();
            compiler.CompilePowerPointMacroFile(outputStream, vbaProjectStream, documentType);

            outputStream.Position = 0;
            return outputStream;
        }

        [Test]
        public void CompilePowerPointAddin_ShouldProduceValidOpenXmlPackage()
        {
            // Arrange
            var (modulePath, classPath) = GetTestDataPaths();
            var compiler = CreateTestCompiler("TestAddin");
            compiler.AddModule(modulePath);
            compiler.AddClass(classPath);

            // Act
            using var addinStream = CompilePowerPointFile(compiler, PresentationDocumentType.AddIn);
            using var presentation = PresentationDocument.Open(addinStream, false);

            // Assert
            ClassicAssert.AreEqual(PresentationDocumentType.AddIn, presentation.DocumentType,
                "Generated file should be a PowerPoint add-in");

            ClassicAssert.IsNotNull(presentation.PresentationPart,
                "Add-in should have a PresentationPart");
        }

        [Test]
        public void CompilePowerPointAddin_ShouldContainVbaProjectPart()
        {
            // Arrange
            var (modulePath, classPath) = GetTestDataPaths();
            var compiler = CreateTestCompiler("TestAddin");
            compiler.AddModule(modulePath);
            compiler.AddClass(classPath);

            // Act
            using var addinStream = CompilePowerPointFile(compiler, PresentationDocumentType.AddIn);
            using var presentation = PresentationDocument.Open(addinStream, false);

            // Assert
            ClassicAssert.IsNotNull(presentation.PresentationPart?.VbaProjectPart,
                "Add-in should have a VbaProjectPart containing the VBA macros");

            var vbaStream = presentation.PresentationPart!.VbaProjectPart!.GetStream();
            ClassicAssert.Greater(vbaStream.Length, 0,
                "VbaProjectPart stream should contain non-zero bytes");
        }

        [Test]
        public void CompilePowerPointMacroPresentation_ShouldProduceValidOpenXmlPackage()
        {
            // Arrange
            var (modulePath, classPath) = GetTestDataPaths();
            var compiler = CreateTestCompiler("TestPresentation");
            compiler.AddModule(modulePath);
            compiler.AddClass(classPath);

            // Act
            using var presentationStream = CompilePowerPointFile(compiler, PresentationDocumentType.MacroEnabledPresentation);
            using var presentation = PresentationDocument.Open(presentationStream, false);

            // Assert
            ClassicAssert.AreEqual(PresentationDocumentType.MacroEnabledPresentation, presentation.DocumentType,
                "Generated file should be a macro-enabled presentation");

            ClassicAssert.IsNotNull(presentation.PresentationPart?.VbaProjectPart,
                "Macro-enabled presentation should have a VbaProjectPart");

            var vbaStream = presentation.PresentationPart!.VbaProjectPart!.GetStream();
            ClassicAssert.Greater(vbaStream.Length, 0,
                "VbaProjectPart stream should contain non-zero bytes");
        }

        [Test]
        public void CompilePowerPointAddin_WithMultipleModules_ShouldProduceValidPackage()
        {
            // Arrange
            var (modulePath, classPath) = GetTestDataPaths();
            var compiler = CreateTestCompiler("MultiModuleAddin");
            compiler.AddModule(modulePath);
            compiler.AddClass(classPath);

            // Act
            using var addinStream = CompilePowerPointFile(compiler, PresentationDocumentType.AddIn);
            using var presentation = PresentationDocument.Open(addinStream, false);

            // Assert
            ClassicAssert.AreEqual(PresentationDocumentType.AddIn, presentation.DocumentType);
            ClassicAssert.IsNotNull(presentation.PresentationPart?.VbaProjectPart);

            var vbaStream = presentation.PresentationPart!.VbaProjectPart!.GetStream();
            ClassicAssert.Greater(vbaStream.Length, 0);
        }
    }
}
