using System;

namespace vbamc.Vba
{
    public class ModuleUnit
    {
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
    }
}
