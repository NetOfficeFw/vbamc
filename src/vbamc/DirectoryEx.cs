using System;

namespace vbamc
{
    public static class DirectoryEx
    {
        public static void EnsureDirectory(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }
    }
}
