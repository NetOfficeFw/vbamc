// Copyright 2022 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

using System;
using System.Diagnostics.Metrics;

namespace vbamc.Vba
{
    public class ModuleUnit
    {
        public const string ModuleHeaderTemplate = "Attribute VB_Name = \"{0}\"\r\n\r\n";
        public const string ClassHeaderTemplate =
            "Attribute VB_Name = \"{0}\"\r\n" +
            "Attribute VB_Base = \"0{{FCFB3D2A-A0FA-1068-A738-08002B3371B5}}\"\r\n" +
            "Attribute VB_GlobalNameSpace = False\r\n" +
            "Attribute VB_Creatable = False\r\n" +
            "Attribute VB_PredeclaredId = False\r\n" +
            "Attribute VB_Exposed = False\r\n" +
            "Attribute VB_TemplateDerived = False\r\n" +
            "Attribute VB_Customizable = False\r\n" +
            "\r\n";
        public const string ThisDocumentHeaderTemplate =
            "Attribute VB_Name = \"{0}\"\r\n" +
            "Attribute VB_Base = \"1Normal.ThisDocument\"\r\n" +
            "Attribute VB_GlobalNameSpace = False\r\n" +
            "Attribute VB_Creatable = False\r\n" +
            "Attribute VB_PredeclaredId = True\r\n" +
            "Attribute VB_Exposed = True\r\n" +
            "Attribute VB_TemplateDerived = True\r\n" +
            "Attribute VB_Customizable = True\r\n" +
            "\r\n";

        private ModuleUnit()
        {}

        public string Name { get; init; } = default!;

        public ModuleUnitType Type { get; init; } = default!;

        public string Content { get; init; } = default!;

        public string NameForProject => this.Type switch
        {
            ModuleUnitType.Document => this.Name + "/&H00000000",
            _ => this.Name,
        };

        public static ModuleUnit FromFile(string path, ModuleUnitType type)
        {
            var name = Path.GetFileNameWithoutExtension(path);
            var content = File.ReadAllText(path);
            var userProfilePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

            content = content.Replace("~/", userProfilePath + Path.DirectorySeparatorChar);

            var module = new ModuleUnit
            {
                Name = name,
                Type = type,
                Content = content
            };

            return module;
        }

        public string ToModuleCode()
        {
            var template = this.Type switch
            {
                ModuleUnitType.Document => ThisDocumentHeaderTemplate,
                ModuleUnitType.Module => ModuleHeaderTemplate,
                ModuleUnitType.Class => ClassHeaderTemplate,
                _ => throw new ArgumentOutOfRangeException("ModuleUnitType", "ModuleUnitType value is not supported yet.")
            };

            var header = string.Format(template, this.Name);

            return header + this.Content;
        }
    }
}
