// Copyright 2022 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

using System;
namespace vbamc.Vba
{
    public class ProjectWmRecord
    {
        public ProjectWmRecord(ICollection<string> modules)
        {
            this.Modules = modules;
        }

        public ICollection<string> Modules { get; }

        public byte[] Generate()
        {
            using var memory = new MemoryStream();
            using var writer = new BinaryWriter(memory);
            
            foreach (var module in this.Modules)
            {
                var sName = VbaEncodings.Default.GetBytes(module);
                var uName = VbaEncodings.UTF16.GetBytes(module);

                // ModuleName
                writer.Write(sName);
                writer.Write(VbaEncodings.NULL);

                // ModuleNameUnicode
                writer.Write(uName);
                writer.Write(VbaEncodings.NULL);
                writer.Write(VbaEncodings.NULL);
            }

            // Terminator (2-bytes) (0x0000)
            writer.Write(VbaEncodings.NULL);
            writer.Write(VbaEncodings.NULL);

            writer.Close();
            return memory.ToArray();
        }
    }
}
