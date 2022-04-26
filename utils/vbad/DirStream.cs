using System;
using System.Text;

namespace vbad
{
    public static class DirStream
    {
        public const short SysKindId = 0x0001;
        public const short LcidId = 0x0002;
        public const short LcidInvokeId = 0x0014;
        public const short CodePageId = 0x0003;
        public const short NameId = 0x0004;
        public const short DocStringId = 0x0005;
        public const short DocStringReserved = 0x0040;
        public const short HelpFilePathId = 0x0006;
        public const short HelpFilePathReserved = 0x003D;
        public const short HelpContextId = 0x0007;
        public const short LibFlagsId = 0x0008;
        public const short VersionId = 0x0009;
        public const short ConstantsId = 0x000C;
        public const short ConstantsReserved = 0x003C;

        public const short ReferenceNameId = 0x0016;
        public const short ReferenceProjectId = 0x000E;
        public const short ReferenceNameReserved = 0x003E;
        public const short ReferenceRegisteredId = 0x000D;
        
        public const short ModuleId = 0x000F;
        public const short ModuleCookieId = 0x0013;
        public const short ModuleNameId = 0x0019;
        public const short ModuleStreamNameId = 0x001A;
        public const short ModuleOffsetId = 0x0031;
        public const short ModuleTerminatorId = 0x002B;

        public static IEnumerable<ModuleInfo> GetModules(byte[] dir)
        {
            var reader = new BinaryReader(new MemoryStream(dir));

            // Information record
            reader.ReadRecord(SysKindId);
            reader.ReadRecord(LcidId);
            reader.ReadRecord(LcidInvokeId);
            reader.ReadRecord(CodePageId);
            reader.ReadRecord(NameId);
            reader.ReadRecord(DocStringId);
            reader.ReadRecord(DocStringReserved);
            reader.ReadRecord(HelpFilePathId);
            reader.ReadRecord(HelpFilePathReserved);
            reader.ReadRecord(HelpContextId);
            reader.ReadRecord(LibFlagsId);
            reader.ReadProjectVersionRecord(VersionId);
            reader.ReadRecord(ConstantsId);
            reader.ReadRecord(ConstantsReserved);

            // References record
            while (reader.ReadReferences()) { }

            // Modules record
            var modules = reader.ReadModules().ToList();

            return modules;
        }

        public static void ReadRecord(this BinaryReader reader, short id)
        {
            var idValue = reader.ReadInt16();
            if (idValue != id)
            {
                throw new InvalidOperationException($"Reading record 0x{idValue:X4} when expected record is 0x{id:X4}");
            }

            uint size = reader.ReadUInt32();
            reader.ReadBytes((int)size);
        }
        
        public static void ReadProjectVersionRecord(this BinaryReader reader, short id)
        {
            var idValue = reader.ReadInt16();
            if (idValue != id)
            {
                throw new InvalidOperationException($"Reading record 0x{idValue:X4} when expected record is 0x{id:X4}");
            }

            // size is ignored
            uint size = reader.ReadUInt32();
            reader.ReadBytes(4);
            reader.ReadBytes(2);
        }

        public static bool ReadReferences(this BinaryReader reader)
        {
            var idValue = reader.ReadInt16();
            if (idValue != ReferenceNameId && idValue != ReferenceProjectId)
            {
                reader.BaseStream.Position -= 2;
                return false;
            }

            var nameSize = reader.ReadInt32();
            reader.ReadBytes(nameSize);

            _ = reader.ReadInt16();
            var nameUnicodeSize = reader.ReadInt32();
            reader.ReadBytes(nameUnicodeSize);

            var referenceId = reader.ReadInt16();
            switch (referenceId)
            {
                case ReferenceRegisteredId:
                    _ = reader.ReadInt32();
                    var libSize = reader.ReadInt32();
                    reader.ReadBytes(libSize);
                    _ = reader.ReadInt32();
                    _ = reader.ReadInt16();
                    return true;

                case ReferenceProjectId:
                    var size = reader.ReadInt32();
                    reader.ReadBytes(size);
                    return true;

                default:
                    throw new InvalidOperationException($"Reading reference 0x{referenceId:X4} when expected reference is 0x{ReferenceRegisteredId:X4}");
            }
        }

        public static IEnumerable<ModuleInfo> ReadModules(this BinaryReader reader)
        {
            // Modules record
            var idValue = reader.ReadInt16();
            if (idValue != ModuleId)
            {
                throw new InvalidOperationException($"Reading module 0x{idValue:X4} when expected module is 0x{ModuleId:X4}");
            }

            var size = reader.ReadUInt32();
            var count = reader.ReadUInt16();
            var cookieId = reader.ReadInt16();
            if (cookieId != ModuleCookieId)
            {
                throw new InvalidOperationException($"Reading module cookie 0x{cookieId:X4} when expected module cookie is 0x{ModuleCookieId:X4}");
            }

            var cookieSize = reader.ReadInt32();
            reader.ReadBytes(cookieSize);

            // Module records
            for (int i = 0; i < count; i++)
            {
                var module = reader.ReadModule();
                yield return module;
            }
        }
        public static ModuleInfo ReadModule(this BinaryReader reader)
        {
            // Module record
            var module = new ModuleInfo();

            do
            {
                var id = reader.ReadInt16();
                Console.WriteLine($"Module record id: 0x{id:X4}");
                
                if (id == ModuleStreamNameId)
                {
                    var size = reader.ReadInt32();
                    var nameBytes = reader.ReadBytes(size);
                    var name = Encoding.GetEncoding(1252).GetString(nameBytes);

                    module.Name = name;
                }
                else if (id == ModuleOffsetId)
                {
                    var size = reader.ReadInt32();
                    var offset = reader.ReadUInt32();

                    module.Offset = offset;
                }
                else if (id == ModuleTerminatorId)
                {
                    Console.WriteLine();
                    _ = reader.ReadInt32();
                    Console.WriteLine($"Module {module.Name} at offset 0x{module.Offset:X8}");
                    return module;
                }
                else
                {
                    var size = reader.ReadInt32();
                    reader.ReadBytes(size);
                }

            } while (true);

        }
    }
}
