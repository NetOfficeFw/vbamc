// Copyright 2022 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

using System;

namespace vbamc.Vba
{
    public class DirStream
    {
        public const ushort TerminatorValue = 0x0010;
        public const uint ReservedValue = 0;

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

            // Records
            information.WriteTo(writer);
            references.WriteTo(writer);
            modules.WriteTo(writer);

            // Terminator
            writer.Write(TerminatorValue);
            writer.Write(ReservedValue);

            writer.Close();

            return memory.ToArray();
        }
    }
}
