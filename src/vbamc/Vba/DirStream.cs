using System;

namespace vbamc.Vba
{
    public class DirStream
    {
        public byte[] GetData(ProjectRecord project)
        {
            var information = new InformationRecord();
            information.SysKind = SysKind.Win32;
            information.ProjectName = project.Name;
            information.ProjectVersion = Constants.VersionOffice365;

            var references = new ReferencesRecord();
            var modules = new ModulesRecord(project.Modules);

            var memory = new MemoryStream(1024);
            var writer = new BinaryWriter(memory);

            information.WriteTo(writer);
            references.WriteTo(writer);
            modules.WriteTo(writer);

            writer.Close();

            return memory.ToArray();
        }
    }
}
