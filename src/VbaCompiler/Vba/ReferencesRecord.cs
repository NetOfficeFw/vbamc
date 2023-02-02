// Copyright 2022 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

namespace vbamc.Vba
{
    public class ReferencesRecord
    {
        public void WriteTo(BinaryWriter writer)
        {
            var stdole = new ReferenceRegisteredRecord
            {
                ReferenceName = "stdole",
                Libid = @"*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\SysWOW64\stdole2.tlb#OLE Automation"
            };

            // TODO: change to ProjectReferenceRecord
            var wordNormalDocument = new ReferenceProjectRecord
            {
                ReferenceName = "Normal",
                LibidAbsolute = @"*\CNormal",
                LibidRelative = @"*\CNormal"
            };

            var office = new ReferenceRegisteredRecord
            {
                ReferenceName = "Office",
                Libid = @"*\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16\MSO.DLL#Microsoft Office 16.0 Object Library"
            };

            // PROJECTREFERENCES record
            stdole.WriteTo(writer);

            // TODO: fix
            // if (project.IsTargetWordMacro)
            // wordNormalDocument.WriteTo(writer);

            office.WriteTo(writer);
        }
    }
}
