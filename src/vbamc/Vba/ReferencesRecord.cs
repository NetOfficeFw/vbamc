using System;

namespace vbamc.Vba
{
    public class ReferencesRecord
    {
        public void WriteTo(BinaryWriter writer)
        {
            var stdole = new ReferenceRecord
            {
                ReferenceName = "stdole",
                Libid = @"*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\SysWOW64\stdole2.tlb#OLE Automation"
            };

            var office = new ReferenceRecord
            {
                ReferenceName = "Office",
                Libid = @"*\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16\MSO.DLL#Microsoft Office 16.0 Object Library"
            };

            // PROJECTREFERENCES record
            stdole.WriteTo(writer);
            office.WriteTo(writer);
        }
    }
}
