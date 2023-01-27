// Copyright 2022 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

namespace vbamc.Vba
{
    public class ReferenceRecord
    {
        public const short ReferenceNameId = 0x0016;
        public const short ReferenceNameReserved = 0x003E;

        protected ReferenceRecord(ReferenceRecordType type)
        {
            this.Type = type;
        }

        public ReferenceRecordType Type { get; }

        public string ReferenceName { get; set; } = "";

        public void WriteTo(BinaryWriter writer)
        {
            // REFERENCENAME record
            var nameBytes = VbaEncodings.Default.GetBytes(this.ReferenceName);
            var nameUnicodeBytes = VbaEncodings.UTF16.GetBytes(this.ReferenceName);

            Guard.EnsureNoNullCharacters(nameBytes, nameof(ReferenceName), "Reference name cannot contain null characters.");

            writer.Write(ReferenceNameId);
            writer.Write(nameBytes.Length);
            writer.Write(nameBytes);
            writer.Write(ReferenceNameReserved);
            writer.Write(nameUnicodeBytes.Length);
            writer.Write(nameUnicodeBytes);

            // Reference record: CONTROL / ORIGINAL / REGISTERED / PROJECT
            // For purposes of this implementation, we're only interested in the REGISTERED record.

            // REFERENCEREGISTERED record
            this.WriteInternalTo(writer);
        }

        protected virtual void WriteInternalTo(BinaryWriter writer)
        {
        }
    }
}
