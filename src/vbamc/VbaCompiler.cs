// Copyright 2022 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.CustomUI;
using DocumentFormat.OpenXml.Packaging;
using Kavod.Vba.Compression;
using OpenMcdf;
using vbamc.Vba;

namespace vbamc
{
    public class VbaCompiler
    {
        private IList<ModuleUnit> modules = new List<ModuleUnit>();

        public Guid ProjectId { get; set; }

        public string ProjectName { get; set; } = "Project";

        public string? CompanyName { get; set; }

        public void AddModule(string path)
        {
            var module = ModuleUnit.FromFile(path, ModuleUnitType.Module);
            this.modules.Add(module);
        }

        public void AddClass(string path)
        {
            var @class = ModuleUnit.FromFile(path, ModuleUnitType.Class);
            this.modules.Add(@class);
        }

        public void AddThisDocument(string path)
        {
            var document = ModuleUnit.FromFile(path, ModuleUnitType.Document);
            this.modules.Add(document);
        }

        public string CompileVbaProject(string intermediatePath, string projectFilename)
        {
            var moduleNames = this.modules.OrderBy(m => m.Type).Select(m => m.Name).ToList();

            var storage = new CompoundFile();
            var projectId = this.ProjectId.ToString("B").ToUpperInvariant();

            // PROJECT stream
            var projectStream = storage.RootStorage.AddStream(StreamId.Project);
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
            projectStream.SetData(projectContent);

            // PROJECTwm stream
            var projectWmStream = storage.RootStorage.AddStream(StreamId.ProjectWm);
            var projectWm = new ProjectWmRecord(moduleNames);
            var projectWmContent = projectWm.Generate();
            projectWmStream.SetData(projectWmContent);

            // VBA storage
            var vbaStorage = storage.RootStorage.AddStorage(StorageId.VBA);

            // _VBA_PROJECT stream
            var vbaProjectStream = vbaStorage.AddStream(StreamId.VbaProject);
            var vbaProject = new VbaProjectStream();
            var vbaProjectContent = vbaProject.Generate();
            vbaProjectStream.SetData(vbaProjectContent);

            // dir stream
            var dirStream = vbaStorage.AddStream(StreamId.Dir);
            var dir = new DirStream();
            var dirContent = dir.GetData(project);
            var compressed = VbaCompression.Compress(dirContent);
            dirStream.SetData(compressed);

            // module streams
            foreach (var module in this.modules)
            {
                var moduleStream = new ModuleStream(module);
                moduleStream.WriteTo(vbaStorage);
            }

            DirectoryEx.EnsureDirectory(intermediatePath);

            var dirDebugPath = Path.Combine(intermediatePath, "dir.bin");
            File.WriteAllBytes(dirDebugPath, dirContent);

            var projectOutputPath = Path.Combine(intermediatePath, projectFilename);
            storage.Save(projectOutputPath);

            return projectOutputPath;
        }

        public void CompilePowerPointMacroFile(string outputPath, string baseFilename, string vbaProjectPath, PresentationDocumentType documentType)
        {
            DirectoryEx.EnsureDirectory(outputPath);
            var suffix = documentType == PresentationDocumentType.AddIn ? "Addin.ppam" : "Macro.pptm";
            string outputFileName = baseFilename + suffix;

            var macroTemplatePath = Path.Combine(AppContext.BaseDirectory, @"data/MacroTemplate.potm");
            var macroTemplate = PresentationDocument.CreateFromTemplate(macroTemplatePath);
            var mainDoc = macroTemplate.PresentationPart;
            if (mainDoc != null)
            {
                var vbaProject = mainDoc.AddNewPart<VbaProjectPart>();
                using var reader = File.OpenRead(vbaProjectPath);
                vbaProject.FeedData(reader);
                reader.Close();
            }

            var ribbonPart = macroTemplate.RibbonAndBackstageCustomizationsPart ?? macroTemplate.AddRibbonAndBackstageCustomizationsPart();
            AttachRibbonCustomization(ribbonPart, Directory.GetCurrentDirectory());

            macroTemplate.PackageProperties.Title = this.ProjectName;
            var propCompany = macroTemplate.ExtendedFilePropertiesPart?.Properties.Company;
            if (propCompany != null && !string.IsNullOrEmpty(this.CompanyName))
            {
                propCompany.Text = this.CompanyName;
            }

            macroTemplate.ChangeDocumentType(documentType);
            var targetMacroPath = Path.Combine(outputPath, outputFileName);
            macroTemplate.SaveAs(targetMacroPath);
        }

        public void CompileExcelMacroFile(string outputPath, string baseFilename, string vbaProjectPath, SpreadsheetDocumentType documentType)
        {
            DirectoryEx.EnsureDirectory(outputPath);
            var suffix = documentType == SpreadsheetDocumentType.AddIn ? "Addin.xlam" : "Macro.xlsm";
            string outputFileName = baseFilename + suffix;

            var macroTemplatePath = Path.Combine(AppContext.BaseDirectory, @"data/MacroTemplate.xltx");
            var macroTemplate = SpreadsheetDocument.CreateFromTemplate(macroTemplatePath);
            var mainDoc = macroTemplate.WorkbookPart;
            if (mainDoc != null)
            {
                var vbaProject = mainDoc.AddNewPart<VbaProjectPart>();
                using var reader = File.OpenRead(vbaProjectPath);
                vbaProject.FeedData(reader);
                reader.Close();
            }

            var ribbonPart = macroTemplate.RibbonAndBackstageCustomizationsPart ?? macroTemplate.AddRibbonAndBackstageCustomizationsPart();
            AttachRibbonCustomization(ribbonPart, Directory.GetCurrentDirectory());

            macroTemplate.PackageProperties.Title = this.ProjectName;
            var propCompany = macroTemplate.ExtendedFilePropertiesPart?.Properties.Company;
            if (propCompany != null && !string.IsNullOrEmpty(this.CompanyName))
            {
                propCompany.Text = this.CompanyName;
            }

            macroTemplate.ChangeDocumentType(documentType);
            var targetMacroPath = Path.Combine(outputPath, outputFileName);
            macroTemplate.SaveAs(targetMacroPath);
        }

        public void CompileWordMacroFile(string outputPath, string baseFilename, string vbaProjectPath, WordprocessingDocumentType documentType)
        {
            if (documentType != WordprocessingDocumentType.MacroEnabledDocument)
            {
                throw new ArgumentOutOfRangeException(nameof(documentType), "Compiler supports only WordprocessingDocumentType.MacroEnabledDocument value.");
            }

            DirectoryEx.EnsureDirectory(outputPath);
            var suffix = "Macro.docm";
            string outputFileName = baseFilename + suffix;

            var macroTemplatePath = Path.Combine(AppContext.BaseDirectory, @"data/MacroTemplate.dotx");
            var macroTemplate = WordprocessingDocument.CreateFromTemplate(macroTemplatePath);
            var mainDoc = macroTemplate.MainDocumentPart;
            if (mainDoc != null)
            {
                var vbaProject = mainDoc.AddNewPart<VbaProjectPart>();
                using var reader = File.OpenRead(vbaProjectPath);
                vbaProject.FeedData(reader);
                reader.Close();
            }

            var ribbonPart = macroTemplate.RibbonAndBackstageCustomizationsPart ?? macroTemplate.AddRibbonAndBackstageCustomizationsPart();
            AttachRibbonCustomization(ribbonPart, Directory.GetCurrentDirectory());

            macroTemplate.PackageProperties.Title = this.ProjectName;
            var propCompany = macroTemplate.ExtendedFilePropertiesPart?.Properties.Company;
            if (propCompany != null && !string.IsNullOrEmpty(this.CompanyName))
            {
                propCompany.Text = this.CompanyName;
            }

            macroTemplate.ChangeDocumentType(documentType);
            var targetMacroPath = Path.Combine(outputPath, outputFileName);
            macroTemplate.SaveAs(targetMacroPath);
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
}
