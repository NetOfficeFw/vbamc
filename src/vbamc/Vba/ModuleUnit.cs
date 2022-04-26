using System;

namespace vbamc.Vba
{
    public class ModuleUnit
    {
        public const string ModuleHeaderTemplate = "Attribute VB_Name = \"{0}\"\r\n\r\n";
        public const string ModuleClassTemplate =
            "Attribute VB_Name = \"{0}\"\r\n" +
            "Attribute VB_Base = \"0{{FCFB3D2A-A0FA-1068-A738-08002B3371B5}}\"\r\n" +
            "Attribute VB_GlobalNameSpace = False\r\n" +
            "Attribute VB_Creatable = False\r\n" +
            "Attribute VB_PredeclaredId = False\r\n" +
            "Attribute VB_Exposed = False\r\n" +
            "Attribute VB_TemplateDerived = False\r\n" +
            "Attribute VB_Customizable = False\r\n" +
            "\r\n";

        public string Name { get; set; }

        public ModuleUnitType Type { get; set; }

        public string Content { get; set; }

        public static ModuleUnit FromFile(string path, ModuleUnitType type)
        {
            var name = Path.GetFileNameWithoutExtension(path);
            var content = File.ReadAllText(path);

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
            var template = this.Type == ModuleUnitType.Module ? ModuleHeaderTemplate : ModuleClassTemplate;
            var header = string.Format(template, this.Name);

            return header + this.Content;
        }
    }
}
