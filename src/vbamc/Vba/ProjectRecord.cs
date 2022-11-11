// Copyright 2022 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

using System;

namespace vbamc.Vba
{
    public class ProjectRecord
    {
        public ProjectRecord()
        {
            this.HostExtenders = new List<string>()
            {
                Constants.HostExtender_VBE
            };
        }

        public string Id { get; set; }

        public string Name { get; set; }

        public string VersionCompatible32 => Constants.VersionCompatible32;

        /// <summary>
        /// CMG value.
        /// </summary>
        public string ProtectionState { get; set; }
        
        /// <summary>
        /// DPB value.
        /// </summary>
        public string ProjectPassword { get; set; }
        
        /// <summary>
        /// GC value.
        /// </summary>
        public string VisibilityState { get; set; }

        public IList<string> HostExtenders { get; }

        public ICollection<ModuleUnit> Modules { get; set; }

        public byte[] Generate()
        {
            var memory = new MemoryStream();
            var writer = new BinaryWriter(memory);
            
            writer.WriteLine($@"ID=""{this.Id}""");

            foreach (var module in this.Modules)
            {
                writer.WriteLine($"{module.Type}={module.Name}");
            }

            writer.WriteLine($@"Name=""{this.Name}""");
            writer.WriteLine($@"HelpContextID=""0""");
            writer.WriteLine($@"VersionCompatible32=""{this.VersionCompatible32}""");
            writer.WriteLine($@"CMG=""{this.ProtectionState}""");
            writer.WriteLine($@"DPB=""{this.ProjectPassword}""");
            writer.WriteLine($@"GC=""{this.VisibilityState}""");
            writer.WriteLine();

            // HostExtenders
            writer.WriteLine($@"[Host Extender Info]");
            foreach (var hostExtender in this.HostExtenders)
            {
                writer.WriteLine(hostExtender);
            }

            writer.Close();
            return memory.ToArray();
        }
    }
}
