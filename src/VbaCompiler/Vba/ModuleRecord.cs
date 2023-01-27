// Copyright 2022 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

using System;

namespace vbamc.Vba
{
    public class ModuleRecord
    {
        public const short NameId = 0x0019;
        public const short NameUnicodeId = 0x0047;
        public const short StreamNameId = 0x001A;
        public const short StreamNameReserved = 0x0032;
        public const short DocStringId = 0x001C;
        public const short DocStringReserved = 0x0048;
        public const short OffsetId = 0x0031;
        public const short HelpContextId = 0x001E;
        public const short CookieId = 0x002C;
        public const short TypeModuleId = 0x0021;
        public const short TypeClassId = 0x0022;
        public const short TypeDocumentId = 0x0022;
        public const short ReadOnlyId = 0x0025;
        public const short PrivateId = 0x0028;

        public const ushort CookieValue = 0xFFFF;
        public const ushort TerminatorValue = 0x002B;
        public const uint ReservedValue = 0;

        /// <summary>
        /// Offset to the module data in the module stream.
        /// </summary>
        /// <remarks>
        /// We always store module at the offset 0 as we do not generate
        /// any performance cache data for the module.
        /// </remarks>
        public const uint ModuleOffset = 0;

        public ModuleRecord(ModuleUnit module)
        {
            this.Module = module;
        }

        public ModuleUnit Module { get; }

        public string DocString => "";

        public int HelpContext => 0;

        public void WriteTo(BinaryWriter writer)
        {
            // MODULENAME record
            var nameBytes = VbaEncodings.Default.GetBytes(this.Module.Name);
            Guard.EnsureNoNullCharacters(nameBytes, "ModuleName", "Module name cannot contain null characters.");

            writer.Write(NameId);
            writer.Write(nameBytes.Length);
            writer.Write(nameBytes);

            // MODULENAMEUNICODE record
            var nameUnicodeBytes = VbaEncodings.UTF16.GetBytes(this.Module.Name);

            writer.Write(NameUnicodeId);
            writer.Write(nameUnicodeBytes.Length);
            writer.Write(nameUnicodeBytes);

            // MODULESTREAMNAME record
            writer.Write(StreamNameId);
            writer.Write(nameBytes.Length);
            writer.Write(nameBytes);
            writer.Write(StreamNameReserved);
            writer.Write(nameUnicodeBytes.Length);
            writer.Write(nameUnicodeBytes);

            // MODULEDOCSTRING record
            var docStringBytes = VbaEncodings.Default.GetBytes(this.DocString);
            var docStringUnicodeBytes = VbaEncodings.UTF16.GetBytes(this.DocString);

            writer.Write(DocStringId);
            writer.Write(docStringBytes.Length);
            writer.Write(docStringBytes);
            writer.Write(DocStringReserved);
            writer.Write(docStringUnicodeBytes.Length);
            writer.Write(docStringUnicodeBytes);

            // MODULEOFFSET record
            writer.Write(OffsetId);
            writer.Write(sizeof(uint));
            writer.Write(ModuleOffset);

            // MODULEHELPCONTEXT record
            writer.Write(HelpContextId);
            writer.Write(sizeof(uint));
            writer.Write(this.HelpContext);

            // MODULECOOKIE record
            writer.Write(CookieId);
            writer.Write(sizeof(ushort));
            writer.Write(CookieValue);

            // MODULETYPE record
            var typeId = GetModuleTypeId(this.Module.Type);

            writer.Write(typeId);
            writer.Write(ReservedValue);

            // MODULEPRIVATE record
            if (this.Module.Type == ModuleUnitType.Class)
            {
                writer.Write(PrivateId);
                writer.Write(ReservedValue);
            }

            // Terminator
            writer.Write(TerminatorValue);

            // Reserved
            writer.Write(ReservedValue);
        }

        /// <summary>
        /// Section 2.3.4.2.3.2.8 MODULETYPE Record
        ///
        /// MUST be 0x0021 when the containing MODULE Record (section 2.3.4.2.3.2) is a procedural module.
        /// MUST be 0x0022 when the containing MODULE Record (section 2.3.4.2.3.2) is a document module, class module, or designer module.
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        private static short GetModuleTypeId(ModuleUnitType type)
        {
            return type switch
            {
                ModuleUnitType.Document => TypeDocumentId,
                ModuleUnitType.Module => TypeModuleId,
                ModuleUnitType.Class => TypeClassId,
                _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
            };
        }
    }
}
