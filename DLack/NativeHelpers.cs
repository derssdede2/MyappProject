using System;
using System.IO;
using System.Runtime.InteropServices;

namespace DLack
{
    /// <summary>
    /// Shared native helpers used by both Scanner and Optimizer.
    /// </summary>
    internal static class NativeHelpers
    {
        [DllImport("Shell32.dll", CharSet = CharSet.Unicode)]
        private static extern int SHQueryRecycleBin(string pszRootPath, ref SHQUERYRBINFO pSHQueryRBInfo);

        [StructLayout(LayoutKind.Sequential, Pack = 4)]
        private struct SHQUERYRBINFO
        {
            public int cbSize;
            public long i64Size;
            public long i64NumItems;
        }

        /// <summary>Returns the total Recycle Bin size in megabytes across all drives.</summary>
        public static long GetRecycleBinSizeMB()
        {
            try
            {
                var info = new SHQUERYRBINFO { cbSize = Marshal.SizeOf<SHQUERYRBINFO>() };
                if (SHQueryRecycleBin(null, ref info) == 0)
                    return info.i64Size / (1024 * 1024);
            }
            catch { }
            return 0;
        }

        /// <summary>
        /// Returns the total size of all files under <paramref name="path"/> in megabytes.
        /// Returns 0 if the directory does not exist or is inaccessible.
        /// </summary>
        public static long GetFolderSizeMB(string path, int maxDepth = int.MaxValue)
        {
            if (!Directory.Exists(path)) return 0;
            try
            {
                long totalBytes = 0;
                var options = new EnumerationOptions
                {
                    IgnoreInaccessible = true,
                    RecurseSubdirectories = maxDepth > 1,
                    MaxRecursionDepth = maxDepth
                };
                foreach (var file in Directory.EnumerateFiles(path, "*", options))
                {
                    try { totalBytes += new FileInfo(file).Length; }
                    catch { }
                }
                return totalBytes / (1024 * 1024);
            }
            catch { return 0; }
        }
    }
}
