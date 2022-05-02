using System;

namespace vbamc.Vba
{
    public class ProjectProtectionState
    {
        public ProjectProtectionState(string projectId)
        {
            ProjectId = projectId;
        }

        public string ProjectId { get; }

        public ProjectProtection ProjectProtection { get; set;  }

        public string ToEncryptedString()
        {
            byte seed = VbaEncryption.GenerateSeed();
            return ToEncryptedString(seed);
        }

        internal string ToEncryptedString(byte seed)
        {
            var enc = VbaEncryption.Encrypt(seed, this.ProjectId, BitConverter.GetBytes(0));

            return HexEncoder.BytesToHexString(enc);
        }
    }
}
