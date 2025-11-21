// Copyright 2022 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

using Kavod.Vba.Compression;
using OpenMcdf;
using System.Text;
using vbad;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

Console.WriteLine("VBA Project Decompiler");

var sourcePath = @"d:\dev\github\NetOfficeFw\vbamc\data\powerpoint\Macro2003\ppt\";
var targetPath = @"d:\dev\github\NetOfficeFw\vbamc\data\powerpoint\Macro2003_vbaProject";

var source = Path.Combine(sourcePath, "vbaProject.bin");

using var root = RootStorage.OpenRead(source);

ProcessEntriesRecursive(root, "");
var vbaStorage = root.OpenStorage("VBA");
ProcessDirAndModules(vbaStorage);

Console.WriteLine();
Console.WriteLine($"  vbaProject.bin -> {targetPath}");
Console.WriteLine("Project was decompiled.");

void ProcessEntriesRecursive(Storage storage, string directory)
{
    foreach (var entry in storage.EnumerateEntries())
    {
        if (entry.Type == EntryType.Stream)
        {
            var name = entry.Name;
            var target = Path.Combine(targetPath, directory, name + ".bin");

            using var stream = storage.OpenStream(name);
            var data = new byte[stream.Length];
            stream.ReadExactly(data);

            if (name == "dir")
            {
                try
                {
                    data = VbaCompression.Decompress(data);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            
            File.WriteAllBytes(target, data);
        }
        else if (entry.Type == EntryType.Storage)
        {
            var dir = Path.Combine(targetPath, entry.Name);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            var subStorage = storage.OpenStorage(entry.Name);
            ProcessEntriesRecursive(subStorage, entry.Name);
        }
    }
}


void ProcessDirAndModules(Storage vbaStorage)
{
    using var dirStream = vbaStorage.OpenStream("dir");
    var data = new byte[dirStream.Length];
    dirStream.ReadExactly(data);
    data = VbaCompression.Decompress(data);
    var target = Path.Combine(targetPath, "vba", "dir.bin");

    File.WriteAllBytes(target, data);

    var modules = DirStream.GetModules(data);
    foreach (var module in modules)
    {
        var name = module.Name;
        if (name == null) continue;
        
        var targetVbBin = Path.Combine(targetPath, "vba", name + ".bin");
        var targetVbText = Path.Combine(targetPath, "vba", name + ".vb");

        using var stream = vbaStorage.OpenStream(name);
        var dataVb = new byte[stream.Length];
        stream.ReadExactly(dataVb);
        var dataVbBin = dataVb.AsSpan().Slice((int)module.Offset).ToArray();
        File.WriteAllBytes(targetVbBin, dataVbBin);
        
        var decompressed = VbaCompression.Decompress(dataVbBin);

        var sourceCode = Encoding.GetEncoding(1252).GetString(decompressed);

        File.WriteAllText(targetVbText, sourceCode);
    }
}