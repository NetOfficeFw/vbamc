// Copyright 2026 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

using System;
using System.IO;
using System.Linq;
using System.Text;
using Kavod.Vba.Compression;
using OpenMcdf;
using VbadDecompiler = vbad;

namespace vbamc.Tests.Streams
{
    /// <summary>
    /// Integration tests that validate VBA project compilation produces
    /// a valid CompoundFile that can be decompiled using vbad logic.
    ///
    /// These tests ensure the compiled VBA project binary has correct
    /// structure and stream mappings. With broken OpenMcdf versions,
    /// the compilation produces files where streams are not correctly
    /// mapped to folders, making the VBA project unusable.
    /// </summary>
    class VbaProjectRoundTripTests
    {
        private byte[] _compiledVbaProject = null!;
        private byte[] _compiledMultiModuleVbaProject = null!;

        [SetUp]
        public void Setup()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var sourcePath = Path.Combine(TestContext.CurrentContext.TestDirectory, "data");
            var modulePath = Path.Combine(sourcePath, "TestModule.vb");
            var classPath = Path.Combine(sourcePath, "Class.vb");

            // Compile single-module project
            var compiler = CreateTestCompiler("RoundTripTest");
            compiler.AddModule(modulePath);
            using var stream = compiler.CompileVbaProject();
            _compiledVbaProject = stream.ToArray();

            // Compile multi-module project
            var multiCompiler = CreateTestCompiler("MultiModuleTest");
            multiCompiler.AddModule(modulePath);
            multiCompiler.AddClass(classPath);
            using var multiStream = multiCompiler.CompileVbaProject();
            _compiledMultiModuleVbaProject = multiStream.ToArray();
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

        [Test]
        public void CompileVbaProject_ShouldProduceDecompilableCompoundFile()
        {
            using var compoundFile = new CompoundFile(new MemoryStream(_compiledVbaProject));
            var root = compoundFile.RootStorage;

            // Verify PROJECT stream exists and has content
            // Empty stream data indicates directory metadata wasn't flushed (broken OpenMcdf)
            var projectStream = root.GetStream("PROJECT");
            var projectData = projectStream.GetData();
            ClassicAssert.Greater(projectData.Length, 0,
                "PROJECT stream is empty - stream may not have been disposed");

            // Verify PROJECTwm stream exists and has content
            var projectWmStream = root.GetStream("PROJECTwm");
            var projectWmData = projectWmStream.GetData();
            ClassicAssert.Greater(projectWmData.Length, 0,
                "PROJECTwm stream is empty - stream may not have been disposed");

            // Verify VBA storage exists
            var vbaStorage = root.GetStorage("VBA");
            ClassicAssert.IsNotNull(vbaStorage, "VBA storage should exist");

            // Verify _VBA_PROJECT stream exists and has content
            var vbaProjectSubStream = vbaStorage.GetStream("_VBA_PROJECT");
            var vbaProjectData = vbaProjectSubStream.GetData();
            ClassicAssert.Greater(vbaProjectData.Length, 0,
                "_VBA_PROJECT stream is empty - stream may not have been disposed");

            // Verify dir stream exists, has content, and can be decompressed
            var dirStream = vbaStorage.GetStream("dir");
            var dirData = dirStream.GetData();
            ClassicAssert.Greater(dirData.Length, 0,
                "dir stream is empty - stream may not have been disposed");

            var decompressedDir = VbaCompression.Decompress(dirData);
            ClassicAssert.Greater(decompressedDir.Length, 0, "Decompressed dir should have content");
        }

        [Test]
        public void CompileVbaProject_ShouldProduceReadableModules()
        {
            using var compoundFile = new CompoundFile(new MemoryStream(_compiledVbaProject));
            var vbaStorage = compoundFile.RootStorage.GetStorage("VBA");

            // Verify dir stream has content before decompression
            var dirStreamData = vbaStorage.GetStream("dir").GetData();
            ClassicAssert.Greater(dirStreamData.Length, 0,
                "dir stream is empty - stream may not have been disposed");

            var dirData = VbaCompression.Decompress(dirStreamData);

            var modules = VbadDecompiler.DirStream.GetModules(dirData).ToList();
            ClassicAssert.Greater(modules.Count, 0, "Should have at least one module");

            foreach (var module in modules)
            {
                ClassicAssert.IsNotEmpty(module.Name, "Module name should not be empty");

                var moduleStream = vbaStorage.GetStream(module.Name);
                var moduleData = moduleStream.GetData();

                // Check raw stream data is not empty before processing
                ClassicAssert.Greater(moduleData.Length, 0,
                    $"Module '{module.Name}' stream is empty - stream may not have been disposed");

                ClassicAssert.Greater(moduleData.Length, (int)module.Offset,
                    $"Module '{module.Name}' data length ({moduleData.Length}) should be greater than offset ({module.Offset})");

                var moduleCode = moduleData.AsSpan().Slice((int)module.Offset).ToArray();
                ClassicAssert.Greater(moduleCode.Length, 0,
                    $"Module '{module.Name}' code after offset is empty");

                // Decompress module code - this is the critical test
                // If the CompoundFile structure is broken, decompression will fail
                var decompressedCode = VbaCompression.Decompress(moduleCode);
                ClassicAssert.Greater(decompressedCode.Length, 0, "Decompressed code should have content");
            }
        }

        [Test]
        public void CompileVbaProject_DecompiledCodeShouldContainOriginalSource()
        {
            using var compoundFile = new CompoundFile(new MemoryStream(_compiledVbaProject));
            var vbaStorage = compoundFile.RootStorage.GetStorage("VBA");

            // Verify dir stream has content before decompression
            var dirStreamData = vbaStorage.GetStream("dir").GetData();
            ClassicAssert.Greater(dirStreamData.Length, 0,
                "dir stream is empty - stream may not have been disposed");

            var dirData = VbaCompression.Decompress(dirStreamData);

            var modules = VbadDecompiler.DirStream.GetModules(dirData).ToList();
            var testModule = modules.First(m => m.Name == "TestModule");

            var moduleStream = vbaStorage.GetStream(testModule.Name!);
            var moduleData = moduleStream.GetData();

            // Check raw stream data is not empty
            ClassicAssert.Greater(moduleData.Length, 0,
                "TestModule stream is empty - stream may not have been disposed");

            var moduleCode = moduleData.AsSpan().Slice((int)testModule.Offset).ToArray();
            ClassicAssert.Greater(moduleCode.Length, 0,
                "TestModule code after offset is empty");

            var decompressedSource = Encoding.GetEncoding(1252).GetString(VbaCompression.Decompress(moduleCode));

            ClassicAssert.IsTrue(decompressedSource.Contains("HelloWorld"),
                "Decompiled source should contain 'HelloWorld' function");
            ClassicAssert.IsTrue(decompressedSource.Contains("AddNumbers"),
                "Decompiled source should contain 'AddNumbers' function");
            ClassicAssert.IsTrue(decompressedSource.Contains("MsgBox"),
                "Decompiled source should contain 'MsgBox' call");
        }

        [Test]
        public void CompileVbaProject_WithMultipleModules_AllShouldBeDecompilable()
        {
            using var compoundFile = new CompoundFile(new MemoryStream(_compiledMultiModuleVbaProject));
            var vbaStorage = compoundFile.RootStorage.GetStorage("VBA");

            // Verify dir stream has content before decompression
            var dirStreamData = vbaStorage.GetStream("dir").GetData();
            ClassicAssert.Greater(dirStreamData.Length, 0,
                "dir stream is empty - stream may not have been disposed");

            var dirData = VbaCompression.Decompress(dirStreamData);

            var modules = VbadDecompiler.DirStream.GetModules(dirData).ToList();
            ClassicAssert.AreEqual(2, modules.Count, "Should have exactly 2 modules (TestModule and Class)");

            foreach (var module in modules)
            {
                var moduleStream = vbaStorage.GetStream(module.Name);
                var moduleData = moduleStream.GetData();

                // Check raw stream data is not empty before processing
                ClassicAssert.Greater(moduleData.Length, 0,
                    $"Module '{module.Name}' stream is empty - stream may not have been disposed");

                var moduleCode = moduleData.AsSpan().Slice((int)module.Offset).ToArray();
                ClassicAssert.Greater(moduleCode.Length, 0,
                    $"Module '{module.Name}' code after offset is empty");

                var decompressedCode = VbaCompression.Decompress(moduleCode);
                ClassicAssert.IsNotNull(decompressedCode, $"Module '{module.Name}' should decompress successfully");
            }
        }
    }
}
