// Copyright 2022 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

using System;

namespace vbamc.Vba
{
    public class ReferenceRecord
    {
        public const short ReferenceNameId = 0x0016;
        public const short ReferenceNameReserved = 0x003E;

        public string ReferenceName { get; set; } = "";
        
        public string Libid { get; set; } = "";

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
            var libidBytes = VbaEncodings.Default.GetBytes(this.Libid);
            var sizeOfLibid = libidBytes.Length;

            Guard.EnsureNoNullCharacters(libidBytes, nameof(Libid), "Libid cannot contain null characters.");

            int size = sizeof(int) + sizeOfLibid + sizeof(int) + sizeof(short);
            
            writer.Write((ushort)ReferenceRecordType.Registered);
            writer.Write(size);
            writer.Write(sizeOfLibid);
            writer.Write(libidBytes);
            writer.Write((uint)0);
            writer.Write((ushort)0);
        }
    }
}
