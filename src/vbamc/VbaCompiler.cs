using System;
using OpenMcdf;
using vbamc.Vba;

namespace vbamc
{
    public class VbaCompiler
    {
        private IDictionary<string, ModuleUnit> modules = new Dictionary<string, ModuleUnit>();

        public Guid ProjectId { get; set; }

        public string ProjectName { get; set; }

        public void AddModule(string path)
        {
            var module = ModuleUnit.FromFile(path, ModuleUnitType.Module);
            this.modules.Add(module.Name, module);
        }
        
        public void AddClass(string path)
        {
            var @class = ModuleUnit.FromFile(path, ModuleUnitType.Class);
            this.modules.Add(@class.Name, @class);
        }

        public void Compile(string targetPath)
        {
            var moduleNames = this.modules.Keys.OrderBy(k => k);

            var storage = new CompoundFile();

            // PROJECT stream
            var projectStream = storage.RootStorage.AddStream(StreamId.Project);
            var project = new ProjectRecord();
            project.Id = this.ProjectId;
            project.Name = this.ProjectName;
            project.Modules = this.modules.Values.OrderBy(v => v.Name);

            // dummy values
            project.ProtectionState = "5351A4A28AA68AA68AA68AA6";
            project.ProjectPassword = "3D3FCADC5EC75FC75FC7";
            project.VisibilityState = "3634C1D243323D333D33C2";

            var projectContent = project.Generate();
            projectStream.SetData(projectContent);

            // PROJECTwm stream
            var projectWmStream = storage.RootStorage.AddStream(StreamId.ProjectWm);
            var projectWm = new ProjectWmRecord(moduleNames);
            var projectWmContent = projectWm.Generate();
            projectWmStream.SetData(projectWmContent);

            storage.Save(targetPath);
        }
    }
}
