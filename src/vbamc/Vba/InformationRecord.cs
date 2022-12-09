// Copyright 2022 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

using System;

namespace vbamc.Vba
{
    public class InformationRecord
    {
        public const short SysKindId = 0x0001;
        public const short CompatVersionId = 0x004A;
        public const short LcidId = 0x0002;
        public const short LcidInvokeId = 0x0014;
        public const short CodePageId = 0x0003;
        public const short NameId = 0x0004;
        public const short DocStringId = 0x0005;
        public const short DocStringReserved = 0x0040;
        public const short HelpFilePathId = 0x0006;
        public const short HelpFilePathReserved = 0x003D;
        public const short HelpContextId = 0x0007;
        public const short LibFlagsId = 0x0008;
        public const short VersionId = 0x0009;
        public const short ConstantsId = 0x000C;
        public const short ConstantsReserved = 0x003C;

        public const int LcidValue = 0x0409;

        public SysKind SysKind { get; set; }

        public string? ProjectName { get; set; }
        
        public string? ProjectDescription { get; set; }

        public int CompatVersion => 2;

        public int Lcid => LcidValue;

        public int LcidInvoke => LcidValue;
        
        public int CodePage => VbaEncodings.Default.CodePage;

        public string ProjectHelpFilePath => "";
        
        public int  ProjectHelpContext => 0;
        
        /// <summary>
        /// ProjectLibFlags MUST have value 0x00000000.
        /// </summary>
        public int  ProjectLibFlags => 0;

        public Version? ProjectVersion { get; set; }

        public string? ProjectConstants { get; set; }


        public void WriteTo(BinaryWriter writer)
        {
            if (this.ProjectVersion == null)
            {
                this.ProjectVersion = Constants.VersionOffice365;
            }

            this.ProjectName = this.ProjectName ?? "";
            this.ProjectDescription = this.ProjectDescription ?? "";
            this.ProjectConstants = this.ProjectConstants ?? "";
            
            // PROJECTSYSKIND record
            writer.Write(SysKindId);
            writer.Write(4);
            writer.Write((int)this.SysKind);

            // PROJECTCOMPATVERSION record
            writer.Write(CompatVersionId);
            writer.Write(4);
            writer.Write(this.CompatVersion);

            // PROJECTLCID record
            writer.Write(LcidId);
            writer.Write(4);
            writer.Write(this.Lcid);

            // PROJECTLCIDINVOKE record
            writer.Write(LcidInvokeId);
            writer.Write(4);
            writer.Write(this.LcidInvoke);

            // PROJECTCODEPAGE record
            writer.Write(CodePageId);
            writer.Write(2);
            writer.Write((short)this.CodePage);

            // PROJECTNAME record
            var nameBytes = VbaEncodings.Default.GetBytes(this.ProjectName);
            Guard.EnsureLengthBetween(nameBytes, 1, 128, nameof(ProjectName), "Project name length must be between 1 and 128 bytes long.");

            writer.Write(NameId);
            writer.Write(nameBytes.Length);
            writer.Write(nameBytes);

            // PROJECTDOCSTRING record
            var docStringBytes = VbaEncodings.Default.GetBytes(this.ProjectDescription);
            var docStringUnicodeBytes = VbaEncodings.UTF16.GetBytes(this.ProjectDescription);

            Guard.EnsureLengthBetween(docStringBytes, 0, 2000, nameof(ProjectDescription), "Project description length must be up to 2000 bytes long.");
            Guard.EnsureNoNullCharacters(docStringBytes, nameof(ProjectDescription), "Project description cannot contain null characters.");

            writer.Write(DocStringId);
            writer.Write(docStringBytes.Length);
            writer.Write(docStringBytes);
            writer.Write(DocStringReserved);
            writer.Write(docStringUnicodeBytes.Length);
            writer.Write(docStringUnicodeBytes);

            // PROJECTHELPFILEPATH record
            var helpFilePathBytes = VbaEncodings.Default.GetBytes(this.ProjectHelpFilePath);
            var helpFilePathUnicodeBytes = VbaEncodings.UTF16.GetBytes(this.ProjectHelpFilePath);

            Guard.EnsureLengthBetween(helpFilePathBytes, 0, 260, nameof(ProjectHelpFilePath), "Project help file path length must be up to 260 bytes long.");
            Guard.EnsureNoNullCharacters(helpFilePathBytes, nameof(ProjectHelpFilePath), "Project help file path cannot contain null characters.");

            writer.Write(HelpFilePathId);
            writer.Write(helpFilePathBytes.Length);
            writer.Write(HelpFilePathReserved);
            writer.Write(helpFilePathUnicodeBytes.Length);
            writer.Write(helpFilePathUnicodeBytes);

            // PROJECTHELPCONTEXT record
            writer.Write(HelpContextId);
            writer.Write(4);
            writer.Write(this.ProjectHelpContext);

            // PROJECTLIBFLAGS record
            writer.Write(LibFlagsId);
            writer.Write(4);
            writer.Write(this.ProjectLibFlags);

            // PROJECTVERSION record
            writer.Write(VersionId);
            writer.Write(4);
            writer.Write((int)this.ProjectVersion.Major);
            writer.Write((short)this.ProjectVersion.Minor);

            // PROJECTCONSTANTS record
            var constantsBytes = VbaEncodings.Default.GetBytes(this.ProjectConstants);
            var constantsUnicodeBytes = VbaEncodings.UTF16.GetBytes(this.ProjectConstants);

            Guard.EnsureLengthBetween(constantsBytes, 0, 1015, nameof(ProjectConstants), "Project constants length must be up to 1015 bytes long.");
            Guard.EnsureNoNullCharacters(constantsBytes, nameof(ProjectConstants), "Project constants cannot contain null characters.");

            writer.Write(ConstantsId);
            writer.Write(constantsBytes.Length);
            writer.Write(ConstantsReserved);
            writer.Write(constantsUnicodeBytes.Length);
            writer.Write(constantsUnicodeBytes);
            
            // PROJECTREFERENCES record
            // TODO: implement...
        }
    }
}
