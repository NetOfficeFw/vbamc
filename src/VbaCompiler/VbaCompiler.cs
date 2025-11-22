// Copyright 2022 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.CustomUI;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.CustomProperties;
using Kavod.Vba.Compression;
using OpenMcdf;
using vbamc.Vba;

namespace vbamc
{
    public class VbaCompiler
    {
        private List<ModuleUnit> modules = [];

        public Dictionary<string, string> ExtendedProperties { get; set; } = [];

        public Guid ProjectId { get; set; }

        public string ProjectName { get; set; } = "Project";

        public string? ProjectVersion { get; set; }

        public string? CompanyName { get; set; }

        public string? UserProfilePath { get; set; }

        /// <summary>
        /// Set to true for instance of VbaCompiler if generating add-ins on windows but are used on macOS to 
        /// explicitly set directory separator char for macOS - otherwise platform default is used 
        /// </summary>
        public bool CompileForMacOnWindows { get; set; } = false;

        public void AddModule(string path)
        {
            var module = ModuleUnit.FromFile(path, ModuleUnitType.Module, this.UserProfilePath, this.CompileForMacOnWindows);
            this.modules.Add(module);
        }

        public void AddClass(string path)
        {
            var @class = ModuleUnit.FromFile(path, ModuleUnitType.Class, this.UserProfilePath, this.CompileForMacOnWindows);
            this.modules.Add(@class);
        }

        public void AddThisDocument(string path)
        {
            var document = ModuleUnit.FromFile(path, ModuleUnitType.Document, this.UserProfilePath, this.CompileForMacOnWindows);
            this.modules.Add(document);
        }

        public MemoryStream CompileVbaProject()
        {
            var vbaProjectMemory = new MemoryStream();
            this.CompileVbaProject(vbaProjectMemory);
            vbaProjectMemory.Position = 0;

            return vbaProjectMemory;
        }

        public void CompileVbaProject(Stream stream)
        {
            var moduleNames = this.modules.OrderBy(m => m.Type).Select(m => m.Name).ToList();

            using var storage = RootStorage.Create(stream, flags: StorageModeFlags.LeaveOpen);
            var projectId = this.ProjectId.ToString("B").ToUpperInvariant();

            // PROJECT stream
            var projectStream = storage.CreateStream(StreamId.Project);
            var project = new ProjectRecord();
            project.Id = projectId;
            project.Name = this.ProjectName;
            project.Modules = this.modules;

            var protectionState = new ProjectProtectionState(projectId);
            var projectPassword = new ProjectPassword(projectId);
            var visibilityState = new ProjectVisibilityState(projectId);

            project.ProtectionState = protectionState.ToEncryptedString();
            project.ProjectPassword = projectPassword.ToEncryptedString();
            project.VisibilityState = visibilityState.ToEncryptedString();

            var projectContent = project.Generate();
            projectStream.Write(projectContent, 0, projectContent.Length);

            // PROJECTwm stream
            var projectWmStream = storage.CreateStream(StreamId.ProjectWm);
            var projectWm = new ProjectWmRecord(moduleNames);
            var projectWmContent = projectWm.Generate();
            projectWmStream.Write(projectWmContent, 0, projectWmContent.Length);

            // VBA storage
            var vbaStorage = storage.CreateStorage(StorageId.VBA);

            // _VBA_PROJECT stream
            var vbaProjectStream = vbaStorage.CreateStream(StreamId.VbaProject);
            var vbaProject = new VbaProjectStream();
            var vbaProjectContent = vbaProject.Generate();
            vbaProjectStream.Write(vbaProjectContent, 0, vbaProjectContent.Length);

            // dir stream
            var dirStream = vbaStorage.CreateStream(StreamId.Dir);
            var dir = new DirStream();
            var dirContent = dir.GetData(project);
            var compressed = VbaCompression.Compress(dirContent);
            dirStream.Write(compressed, 0, compressed.Length);

            // module streams
            foreach (var module in this.modules)
            {
                var moduleStream = new ModuleStream(module);
                moduleStream.WriteTo(vbaStorage);
            }

        }

        public void CompilePowerPointMacroFile(Stream outputMacroFileStream, Stream vbaProjectStream, PresentationDocumentType documentType, string? customSourcePath = null)
        {
            var macroTemplatePath = Path.Combine(AppContext.BaseDirectory, @"data/MacroTemplate.potm");
            var macroTemplate = PresentationDocument.CreateFromTemplate(macroTemplatePath);
            var mainDoc = macroTemplate.PresentationPart;
            if (mainDoc != null)
            {
                var vbaProject = mainDoc.AddNewPart<VbaProjectPart>();
                vbaProject.FeedData(vbaProjectStream);
            }

            var ribbonPart = macroTemplate.RibbonAndBackstageCustomizationsPart ?? macroTemplate.AddRibbonAndBackstageCustomizationsPart();
            AttachRibbonCustomization(ribbonPart, customSourcePath ?? Directory.GetCurrentDirectory());

            macroTemplate.PackageProperties.Title = this.ProjectName;
            macroTemplate.PackageProperties.Version = this.ProjectVersion;
            var extendedProperties = macroTemplate.ExtendedFilePropertiesPart?.Properties;

            var propCompany = extendedProperties?.Company;
            if (propCompany != null && !string.IsNullOrEmpty(this.CompanyName))
            {
                propCompany.Text = this.CompanyName;
            }

            AddExtendedProperties(extendedProperties);

            macroTemplate.ChangeDocumentType(documentType);
            using var tempMacroFile = macroTemplate.Clone(outputMacroFileStream);
        }

        public void CompileExcelMacroFile(Stream outputMacroFileStream, Stream vbaProjectStream, SpreadsheetDocumentType documentType, string? customSourcePath = null)
        {
            var macroTemplatePath = Path.Combine(AppContext.BaseDirectory, @"data/MacroTemplate.xltx");
            var macroTemplate = SpreadsheetDocument.CreateFromTemplate(macroTemplatePath);
            var mainDoc = macroTemplate.WorkbookPart;
            if (mainDoc != null)
            {
                var vbaProject = mainDoc.AddNewPart<VbaProjectPart>();
                vbaProject.FeedData(vbaProjectStream);
            }

            var ribbonPart = macroTemplate.RibbonAndBackstageCustomizationsPart ?? macroTemplate.AddRibbonAndBackstageCustomizationsPart();
            AttachRibbonCustomization(ribbonPart, customSourcePath ?? Directory.GetCurrentDirectory());

            macroTemplate.PackageProperties.Title = this.ProjectName;
            macroTemplate.PackageProperties.Version = this.ProjectVersion;
            var extendedProperties = macroTemplate.ExtendedFilePropertiesPart?.Properties;

            var propCompany = extendedProperties?.Company;
            if (propCompany != null && !string.IsNullOrEmpty(this.CompanyName))
            {
                propCompany.Text = this.CompanyName;
            }

            AddExtendedProperties(extendedProperties);

            macroTemplate.ChangeDocumentType(documentType);
            using var macroFile = macroTemplate.Clone(outputMacroFileStream);
        }

        public void CompileWordMacroFile(Stream outputMacroFileStream, Stream vbaProjectStream, WordprocessingDocumentType documentType, string? customSourcePath = null)
        {
            if (documentType != WordprocessingDocumentType.MacroEnabledDocument)
            {
                throw new ArgumentOutOfRangeException(nameof(documentType), "Compiler supports only WordprocessingDocumentType.MacroEnabledDocument value.");
            }

            var macroTemplatePath = Path.Combine(AppContext.BaseDirectory, @"data/MacroTemplate.dotx");
            var macroTemplate = WordprocessingDocument.CreateFromTemplate(macroTemplatePath);
            var mainDoc = macroTemplate.MainDocumentPart;
            if (mainDoc != null)
            {
                var vbaProject = mainDoc.AddNewPart<VbaProjectPart>();
                vbaProject.FeedData(vbaProjectStream);
            }

            var ribbonPart = macroTemplate.RibbonAndBackstageCustomizationsPart ?? macroTemplate.AddRibbonAndBackstageCustomizationsPart();
            AttachRibbonCustomization(ribbonPart, customSourcePath ?? Directory.GetCurrentDirectory());

            macroTemplate.PackageProperties.Title = this.ProjectName;
            macroTemplate.PackageProperties.Version = this.ProjectVersion;
            var extendedProperties = macroTemplate.ExtendedFilePropertiesPart?.Properties;
            var propCompany = extendedProperties?.Company;
            if (propCompany != null && !string.IsNullOrEmpty(this.CompanyName))
            {
                propCompany.Text = this.CompanyName;
            }

            AddExtendedProperties(extendedProperties);

            macroTemplate.ChangeDocumentType(documentType);
            using var macroFile = macroTemplate.Clone(outputMacroFileStream);
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
            //Console.WriteLine($"Added ribbon customization from file '{ribbonPath}'");

            var images = Directory.EnumerateFiles(Path.Combine(customUiDir, "images"), "*.png");
            foreach (var imagePath in images)
            {
                var imageFilename = Path.GetFileNameWithoutExtension(imagePath);
                var imagePart = ribbonPart.AddImagePart(ImagePartType.Png, imageFilename);
                using var imageStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read);
                imagePart.FeedData(imageStream);
            }
        }

        private void AddExtendedProperties(DocumentFormat.OpenXml.ExtendedProperties.Properties? extendedProperties)
        {
            foreach (var property in this.ExtendedProperties)
            {
                var newProperty = new CustomDocumentProperty
                {
                    Name = property.Key,
                    InnerXml = property.Value
                };
                extendedProperties?.AppendChild(newProperty);
            }
            extendedProperties?.Save();
        }
    }
}
