// Copyright 2022 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

using System;

namespace vbamc.Vba
{
    public class ModulesRecord
    {
        public const short ModulesId = 0x000F;
        public const short CookieId = 0x0013;

        public const ushort CookieValue = 0xFFFF;

        public ModulesRecord(ICollection<ModuleUnit> modules)
        {
            this.Modules = modules;
        }

        public ICollection<ModuleUnit> Modules { get; }

        public void WriteTo(BinaryWriter writer)
        {
            // PROJECTMODULES record
            ushort modulesCount = (ushort)this.Modules.Count();

            writer.Write(ModulesId);
            writer.Write(sizeof(short));
            writer.Write(modulesCount);

            // PROJECTCOOKIE record
            writer.Write(CookieId);
            writer.Write(sizeof(short));
            writer.Write(CookieValue);

            foreach (var module in this.Modules)
            {
                var moduleRecord = new ModuleRecord(module);
                moduleRecord.WriteTo(writer);
            }
        }
    }
}
