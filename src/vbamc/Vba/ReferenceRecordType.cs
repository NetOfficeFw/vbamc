using System;

namespace vbamc.Vba
{
    public enum ReferenceRecordType : ushort
    {
        Control = 0x002F,
        Original = 0x0033,
        Registered = 0x000D,
        Project = 0x000E
    }
}
