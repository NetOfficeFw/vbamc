using System.Security.Cryptography;

namespace VbaCompiler.Tests
{
    public class VbaSamplesNumberGenerator : RandomNumberGenerator
    {
        public override void GetBytes(byte[] data)
        {
            for (int i = 0; i < data.Length; i++)
            {
                data[i] = 0x07;
            }
        }

        public override void GetNonZeroBytes(byte[] data)
        {
            GetBytes(data);
        }
    }
}