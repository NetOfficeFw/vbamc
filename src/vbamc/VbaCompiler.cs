using System;
using Kavod.Vba.Compression;
using OpenMcdf;
using vbamc.Vba;

namespace vbamc
{
    public class VbaCompiler
    {
        private IList<ModuleUnit> modules = new List<ModuleUnit>();

        public Guid ProjectId { get; set; }

        public string ProjectName { get; set; }

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

        public string Compile(string intermediatePath, string projectFilename)
        {
            var moduleNames = this.modules.Select(m => m.Name).ToList();

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
    }
}
