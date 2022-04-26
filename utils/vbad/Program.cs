using Kavod.Vba.Compression;
using OpenMcdf;
using System.Text;
using vbad;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

Console.WriteLine("VBA Project Decompiler");

var sourcePath = @"d:\dev\github\NetOfficeFw\vbamc\data\powerpoint\Macro2003\ppt\";
var targetPath = @"d:\dev\github\NetOfficeFw\vbamc\data\powerpoint\Macro2003_vbaProject";

var source = Path.Combine(sourcePath, "vbaProject.bin");

var vba = new CompoundFile(source);
var root = vba.RootStorage;

root.VisitEntries(item => ProcessFile(root, "", item), false);
ProcessDirAndModules(root.GetStorage("VBA"));

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

                var modules = DirStream.GetModules(data);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
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


void ProcessDirAndModules(CFStorage vbaStorage)
{
    var dir = vbaStorage.GetStream("dir");
    var data = dir.GetData();
    data = VbaCompression.Decompress(data);
    var target = Path.Combine(targetPath, "vba", "dir.bin");

    File.WriteAllBytes(target, data);

    var modules = DirStream.GetModules(data);
    foreach (var module in modules)
    {
        var name = module.Name;
        var targetVbBin = Path.Combine(targetPath, "vba", name + ".bin");
        var targetVbText = Path.Combine(targetPath, "vba", name + ".vb");

        var stream = vbaStorage.GetStream(name);
        Span<byte> dataVb = stream.GetData();
        var dataVbBin = dataVb.Slice((int)module.Offset).ToArray();
        File.WriteAllBytes(targetVbBin, dataVbBin);
        
        dataVb = VbaCompression.Decompress(dataVbBin);

        var sourceCode = Encoding.GetEncoding(1252).GetString(dataVb);

        File.WriteAllText(targetVbText, sourceCode);
    }
}