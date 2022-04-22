using System;

namespace vbamc.Vba
{
    public static class Constants
    {
        public static readonly string VersionCompatible32 = "393222000";
        
        public static readonly string HostExtender_VBE = "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000";

        public static readonly Version VersionOffice2003 = new Version(0x645E9423, 6);
        public static readonly Version VersionOffice365 = new Version(0x645BE109, 11);
    }
}
