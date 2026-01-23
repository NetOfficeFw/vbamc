// Copyright 2026 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Kavod.Vba.Compression;
using OpenMcdf;
using vbamc.Vba;
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
        private string _testModuleExpectedSource = null!;
        private string _classExpectedSource = null!;

        [SetUp]
        public void Setup()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var sourcePath = Path.Combine(TestContext.CurrentContext.TestDirectory, "data");
            var modulePath = Path.Combine(sourcePath, "TestModule.vb");
            var classPath = Path.Combine(sourcePath, "Class.vb");

            // Generate expected compiled source using ModuleUnit (includes VBA headers)
            var testModule = ModuleUnit.FromFile(modulePath, ModuleUnitType.Module, null);
            var classModule = ModuleUnit.FromFile(classPath, ModuleUnitType.Class, null);
            _testModuleExpectedSource = testModule.ToModuleCode();
            _classExpectedSource = classModule.ToModuleCode();

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
        public void CompileVbaProject_DecompiledCodeShouldMatchOriginalSource()
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

            ClassicAssert.AreEqual(_testModuleExpectedSource, decompressedSource,
                "Decompiled source should exactly match the original source with VBA headers");
        }

        [Test]
        public void CompileVbaProject_WithMultipleModules_AllShouldMatchOriginalSource()
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

            // Map of module names to their expected source
            var expectedSources = new Dictionary<string, string>
            {
                { "TestModule", _testModuleExpectedSource },
                { "Class", _classExpectedSource }
            };

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

                var decompressedSource = Encoding.GetEncoding(1252).GetString(VbaCompression.Decompress(moduleCode));

                ClassicAssert.IsTrue(expectedSources.ContainsKey(module.Name!),
                    $"Unexpected module '{module.Name}' found in compiled project");
                ClassicAssert.AreEqual(expectedSources[module.Name!], decompressedSource,
                    $"Decompiled source for module '{module.Name}' should exactly match the original source with VBA headers");
            }
        }

        [Test]
        public void CompileVbaProject_ShouldMatchGoldenFile()
        {
            // Use fixed values to ensure deterministic output
            var fixedGuid = new Guid("12345678-1234-1234-1234-123456789ABC");
            var originalRandomGenerator = VbaEncryption.CreateRandomGenerator;

            try
            {
                // Replace random generator with deterministic one
                VbaEncryption.CreateRandomGenerator = () => new DeterministicRandomNumberGenerator();

                var sourcePath = Path.Combine(TestContext.CurrentContext.TestDirectory, "data");
                var modulePath = Path.Combine(sourcePath, "TestModule.vb");

                var compiler = new VbaCompiler
                {
                    ProjectId = fixedGuid,
                    ProjectName = "GoldenTest",
                    ProjectVersion = "1.0.0",
                    CompanyName = "TestCompany"
                };
                compiler.AddModule(modulePath);

                using var stream = compiler.CompileVbaProject();
                var compiledBytes = stream.ToArray();

                // Mask out compound file timestamps that vary between runs
                // These are at offsets in the directory entries (varies by file structure)
                MaskCompoundFileTimestamps(compiledBytes);

                var goldenFilePath = Path.Combine(sourcePath, "vbaProject.bin");

                if (!File.Exists(goldenFilePath))
                {
                    // Generate golden file if it doesn't exist
                    File.WriteAllBytes(goldenFilePath, compiledBytes);
                    Assert.Inconclusive($"Golden file created at {goldenFilePath}. Re-run the test to validate.");
                }

                var goldenBytes = File.ReadAllBytes(goldenFilePath);
                MaskCompoundFileTimestamps(goldenBytes);

                ClassicAssert.AreEqual(goldenBytes, compiledBytes,
                    "Compiled VBA project should match the golden file byte-for-byte (excluding timestamps)");
            }
            finally
            {
                // Restore original random generator
                VbaEncryption.CreateRandomGenerator = originalRandomGenerator;
            }
        }

        /// <summary>
        /// Masks out timestamp fields in compound file directory entries.
        /// Compound files have 128-byte directory entries with timestamps at specific offsets.
        /// </summary>
        private static void MaskCompoundFileTimestamps(byte[] data)
        {
            // Compound file directory entries start at sector 0 (offset 512) or later
            // Each entry is 128 bytes with creation time at offset 100 and modification time at offset 108
            // The directory sector location is specified in the header at offset 48

            if (data.Length < 512)
                return;

            // Read directory sector location from header (offset 48, 4 bytes, little-endian)
            int directorySectorIndex = BitConverter.ToInt32(data, 48);
            if (directorySectorIndex < 0)
                return;

            // Sector size is typically 512 bytes for compound files with version 3
            const int sectorSize = 512;
            const int headerSize = 512;
            int directoryOffset = headerSize + (directorySectorIndex * sectorSize);

            // Mask timestamps in all directory entries
            const int entrySize = 128;
            const int creationTimeOffset = 100;
            const int modificationTimeOffset = 108;
            const int timestampSize = 8;

            for (int entryStart = directoryOffset;
                 entryStart + entrySize <= data.Length;
                 entryStart += entrySize)
            {
                // Check if this looks like a valid entry (first byte of name should be non-zero for used entries)
                if (data[entryStart] == 0 && data[entryStart + 1] == 0)
                    break; // End of entries

                // Mask creation time
                for (int i = 0; i < timestampSize; i++)
                    data[entryStart + creationTimeOffset + i] = 0;

                // Mask modification time
                for (int i = 0; i < timestampSize; i++)
                    data[entryStart + modificationTimeOffset + i] = 0;
            }
        }

        /// <summary>
        /// A deterministic random number generator for testing purposes.
        /// Produces a fixed repeatable sequence independent of .NET version.
        /// </summary>
        private sealed class DeterministicRandomNumberGenerator : RandomNumberGenerator
        {
            private int _index;

            public override void GetBytes(byte[] data)
            {
                for (int i = 0; i < data.Length; i++)
                {
                    data[i] = (byte)((_index++ * 17 + 31) % 256);
                }
            }

            public override void GetNonZeroBytes(byte[] data)
            {
                for (int i = 0; i < data.Length; i++)
                {
                    // Generate values 1-255 (never zero)
                    data[i] = (byte)((_index++ * 17 + 31) % 255 + 1);
                }
            }
        }
    }
}
