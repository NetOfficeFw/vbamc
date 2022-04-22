using System;

namespace vbamc.Vba
{
    public class DirStream
    {
        public byte[] GetData(VbaCompiler project)
        {
            var information = new InformationRecord();
            information.SysKind = SysKind.Win32;
            information.ProjectName = project.ProjectName;
            information.ProjectVersion = Constants.VersionOffice365;

            var references = new ReferencesRecord();

            var memory = new MemoryStream(1024);
            var writer = new BinaryWriter(memory);

            information.WriteTo(writer);
            references.WriteTo(writer);

            writer.Close();

            return memory.ToArray();
        }
    }
}
