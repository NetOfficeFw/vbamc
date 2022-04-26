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

        public void WriteTo(CFStorage storage)
        {
            var streamName = this.Module.Name;
            var stream = storage.AddStream(streamName);

            var content = this.Module.ToModuleCode();
            var contentBytes = VbaEncodings.Default.GetBytes(content);
            var compressedBytes = VbaCompression.Compress(contentBytes);

            stream.SetData(compressedBytes);
        }
    }
}
