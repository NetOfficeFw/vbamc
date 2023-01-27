// Copyright 2022 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

using System;
using System.Security.Cryptography;
using System.Text;

namespace vbamc.Vba
{
    public class VbaEncryption
    {
        public const byte EncryptionVersion = 2;

        public static Func<RandomNumberGenerator> CreateRandomGenerator { get; set; } = () => RandomNumberGenerator.Create();

        public static byte GenerateSeed()
        {
            var rnd = CreateRandomGenerator();
            var buffer = new byte[1];
            rnd.GetNonZeroBytes(buffer);
            return buffer[0];
        }

        public static byte GetProjectKey(string projectId)
        {
            var buffer = Encoding.ASCII.GetBytes(projectId);

            byte projectKey = 0;

            foreach (var b in buffer)
            {
                projectKey += b;
            }

            return projectKey;
        }

        public static ReadOnlySpan<byte> Encrypt(byte seed, string projectId, ReadOnlySpan<byte> data)
        {
            var rnd = CreateRandomGenerator();
            var memory = new MemoryStream(1024);
            var writer = new BinaryWriter(memory);

            var dataLength = data.Length;
            byte projectKey = GetProjectKey(projectId);

            byte versionEnc = (byte)(seed ^ EncryptionVersion);
            byte projectKeyEnc = (byte)(seed ^ projectKey);

            byte unencryptedByte = projectKey;
            byte encryptedByte1 = projectKeyEnc;
            byte encryptedByte2 = versionEnc;

            writer.Write(seed);
            writer.Write(versionEnc);
            writer.Write(projectKeyEnc);

            int ignoredLength = (seed & 6) / 2;
            var tempBytes = new byte[ignoredLength];
            rnd.GetNonZeroBytes(tempBytes);

            for (int i = 0; i < ignoredLength; i++)
            {
                byte tmp = tempBytes[i];
                byte byteEnc = (byte)(tmp ^ (encryptedByte2 + unencryptedByte));
                writer.Write(byteEnc);

                encryptedByte2 = encryptedByte1;
                encryptedByte1 = byteEnc;
                unencryptedByte = tmp;
            }

            var dataLengthBytes = BitConverter.GetBytes(dataLength);
            for (int i = 0; i < dataLengthBytes.Length; i++)
            {
                byte tmp = dataLengthBytes[i];
                byte byteEnc = (byte)(tmp ^ (encryptedByte2 + unencryptedByte));
                writer.Write(byteEnc);

                encryptedByte2 = encryptedByte1;
                encryptedByte1 = byteEnc;
                unencryptedByte = tmp;
            }

            for (int i = 0; i < data.Length; i++)
            {
                byte tmp = data[i];
                byte byteEnc = (byte)(tmp ^ (encryptedByte2 + unencryptedByte));
                writer.Write(byteEnc);

                encryptedByte2 = encryptedByte1;
                encryptedByte1 = byteEnc;
                unencryptedByte = tmp;
            }

            writer.Close();
            return memory.ToArray();
        }
    }
}