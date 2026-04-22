// Copyright 2022 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

using System;
using OpenMcdf;
using Kavod.Vba.Compression;

namespace vbamc.Vba
{
    public class ModuleStream
    {
        public ModuleStream(ModuleUnit module)
        {
            this.Module = module;
        }

        public ModuleUnit Module { get; }

        public void WriteTo(Storage storage)
        {
            var streamName = this.Module.Name;
            using var stream = storage.CreateStream(streamName);

            var content = this.Module.ToModuleCode();
            var contentBytes = VbaEncodings.Default.GetBytes(content);
            var compressedBytes = VbaCompression.Compress(contentBytes);

            stream.Write(compressedBytes, 0, compressedBytes.Length);
        }
    }
}
