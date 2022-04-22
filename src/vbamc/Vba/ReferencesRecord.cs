using System;

namespace vbamc.Vba
{
    public class ReferencesRecord
    {
        public const short Terminator = 0x000F;

        public void WriteTo(BinaryWriter writer)
        {
            // PROJECTREFERENCES record

            // Terminator
            writer.Write(Terminator);
        }
    }
}
