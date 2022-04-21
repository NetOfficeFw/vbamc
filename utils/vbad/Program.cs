using Kavod.Vba.Compression;
using OpenMcdf;


Console.WriteLine("VBA Project Decompiler");

var sourcePath = @"d:\dev\github\NetOfficeFw\vbamc\data\powerpoint\Macro2003\ppt\";
var targetPath = @"d:\dev\github\NetOfficeFw\vbamc\data\powerpoint\Macro2003_vbaProject";

var source = Path.Combine(sourcePath, "vbaProject.bin");

var vba = new CompoundFile(source);
var root = vba.RootStorage;

root.VisitEntries(item => ProcessFile(root, "", item), false);

Console.WriteLine();
Console.WriteLine($"  vbaProject.bin -> {targetPath}");
Console.WriteLine("Project was decompiled.");

void ProcessFile(CFStorage storage, string directory, CFItem item)
{
    if (item.IsStream)
    {
        var name = item.Name;
        var target = Path.Combine(targetPath, directory, name + ".bin");

        var stream = storage.GetStream(name);
        var data = stream.GetData();

        if (name == "dir")
        {
            try
            {
                data = VbaCompression.Decompress(data);
            }
            catch (Exception ex)
            {
            }
        }
        
        File.WriteAllBytes(target, data);
    }
    else if (item.IsStorage)
    {
        var s2 = (CFStorage)item;
        var dir = Path.Combine(targetPath, s2.Name);
        if (!Directory.Exists(dir))
        {
            Directory.CreateDirectory(dir);
        }

        s2.VisitEntries(item => ProcessFile(s2, s2.Name, item), false);
    }
}