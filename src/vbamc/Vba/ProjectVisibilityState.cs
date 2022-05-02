using System;

namespace vbamc.Vba
{
    public class ProjectVisibilityState
    {
        public const byte ProjectNotVisibleValue = 0x00;
        public const byte ProjectVisibleValue = 0xFF;

        public ProjectVisibilityState(string projectId)
        {
            ProjectId = projectId;
            this.IsProjectVisible = true;
        }

        public string ProjectId { get; }

        public bool IsProjectVisible { get; set; }

        public string ToEncryptedString()
        {
            byte seed = VbaEncryption.GenerateSeed();
            return ToEncryptedString(seed);
        }

        internal string ToEncryptedString(byte seed)
        {
            var value = this.IsProjectVisible ? ProjectVisibleValue : ProjectNotVisibleValue;
            var data = new byte[] { value };
            
            var enc = VbaEncryption.Encrypt(seed, this.ProjectId, data);

            return HexEncoder.BytesToHexString(enc);
        }
    }
}
