// Copyright 2022 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

using System;

namespace vbamc.Vba
{
    public class ProjectPassword
    {
        public ProjectPassword(string projectId)
        {
            ProjectId = projectId;
        }

        public string ProjectId { get; }
        
        public string ToEncryptedString()
        {
            byte seed = VbaEncryption.GenerateSeed();
            return ToEncryptedString(seed);
        }

        internal string ToEncryptedString(byte seed)
        {
            var data = new byte[] { 0 };

            var enc = VbaEncryption.Encrypt(seed, this.ProjectId, data);

            return HexEncoder.BytesToHexString(enc);
        }
    }
}
