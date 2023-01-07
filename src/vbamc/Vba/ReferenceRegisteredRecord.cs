// Copyright 2022 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

namespace vbamc.Vba
{
    public class ReferenceRegisteredRecord : ReferenceRecord
    {
        public ReferenceRegisteredRecord()
            : base(ReferenceRecordType.Registered)
        {
        }

        public string Libid { get; set; } = "";

        protected override void WriteInternalTo(BinaryWriter writer)
        {
            // REFERENCEREGISTERED record
            var libidBytes = VbaEncodings.Default.GetBytes(this.Libid);
            var sizeOfLibid = libidBytes.Length;

            Guard.EnsureNoNullCharacters(libidBytes, nameof(Libid), "Libid cannot contain null characters.");

            int size = sizeof(int) + sizeOfLibid + sizeof(uint) + sizeof(ushort);

            writer.Write((ushort)this.Type);
            writer.Write(size);
            writer.Write(sizeOfLibid);
            writer.Write(libidBytes);
            writer.Write((uint)0);
            writer.Write((ushort)0);
        }
    }
}
