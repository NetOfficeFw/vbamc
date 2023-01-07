// Copyright 2022 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

namespace vbamc.Vba
{
    public class ReferenceProjectRecord : ReferenceRecord
    {
        public ReferenceProjectRecord()
            : base(ReferenceRecordType.Project)
        {
        }

        public string LibidAbsolute { get; set; } = "";

        public string LibidRelative { get; set; } = "";

        protected override void WriteInternalTo(BinaryWriter writer)
        {
            // REFERENCEPROJECT record
            var libidAbsoluteBytes = VbaEncodings.Default.GetBytes(this.LibidAbsolute);
            var sizeOfLibidAbsolute = libidAbsoluteBytes.Length;
            var libidRelativeBytes = VbaEncodings.Default.GetBytes(this.LibidRelative);
            var sizeOfLibidRelative = libidAbsoluteBytes.Length;

            Guard.EnsureNoNullCharacters(libidAbsoluteBytes, nameof(LibidAbsolute), "LibidAbsolute cannot contain null characters.");
            Guard.EnsureNoNullCharacters(libidRelativeBytes, nameof(LibidRelative), "LibidRelative cannot contain null characters.");

            int size = sizeof(int) + sizeOfLibidAbsolute + sizeof(int) + sizeOfLibidRelative + sizeof(uint) + sizeof(ushort);

            writer.Write((ushort)this.Type);
            writer.Write(size);
            writer.Write(sizeOfLibidAbsolute);
            writer.Write(libidAbsoluteBytes);
            writer.Write(sizeOfLibidRelative);
            writer.Write(libidRelativeBytes);
            writer.Write((uint)1706735520);
            writer.Write((ushort)0x07);
        }
    }
}
