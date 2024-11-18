using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PptxFileSizeAnalyzer
{
    /// <summary>
    /// Provides extension methods for file size conversions.
    /// </summary>
    public static class FileSizeExtensions
    {
        private static readonly string[] SizeSuffixes = { "B", "KB", "MB", "GB", "TB", "PB", "EB" };

        /// <summary>
        /// Converts a file size in bytes to a human-readable string representation.
        /// </summary>
        /// <param name="bytes">The size of the file in bytes.</param>
        /// <param name="decimalPlaces">The number of decimal places to include in the formatted string.</param>
        /// <returns>A string representing the file size in a human-readable format.</returns>
        public static string ToReadableSize(this long bytes, int decimalPlaces = 2)
        {
            if (bytes < 0) { return "-" + ToReadableSize(-bytes); }
            if (bytes == 0) { return $"0 {SizeSuffixes[0]}"; }

            int mag = (int)Math.Log(bytes, 1024);
            decimal adjustedSize = (decimal)bytes / (1L << (mag * 10));

            return string.Format($"{{0:F{decimalPlaces}}} {{1}}", adjustedSize, SizeSuffixes[mag]);
        }
    }
}
