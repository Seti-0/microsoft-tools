using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.IO;

namespace Red.Core.IO
{
    public class PathHelper
    {
        public static string GetUniqueFileName(string baseName)
        {
            baseName = CreateValidFilename(baseName);

            string extension = Path.GetExtension(baseName);
            baseName = Path.GetFileNameWithoutExtension(baseName);

            string result = baseName + extension;
            int index = 1;

            while (File.Exists(baseName))
            {
                result = $"{baseName} ({index}){extension}";
                index++;
            }

            return result;
        }

        /// <summary>
        /// Strip illegal chars and reserved words from a candidate filename (should not include the directory path)
        /// </summary>
        /// <remarks>
        /// http://stackoverflow.com/questions/309485/c-sharp-sanitize-file-name
        /// </remarks>
        public static string CreateValidFilename(string filename)
        {
            var invalidChars = Regex.Escape(new string(Path.GetInvalidFileNameChars()));
            var invalidReStr = string.Format(@"[{0}]+", invalidChars);

            var reservedWords = new[]
            {
                "CON", "PRN", "AUX", "CLOCK$", "NUL", "COM0", "COM1", "COM2", "COM3", "COM4",
                "COM5", "COM6", "COM7", "COM8", "COM9", "LPT0", "LPT1", "LPT2", "LPT3", "LPT4",
                "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
            };

            var sanitisedNamePart = Regex.Replace(filename, invalidReStr, "_");
            foreach (var reservedWord in reservedWords)
            {
                var reservedWordPattern = string.Format("^{0}\\.", reservedWord);
                sanitisedNamePart = Regex.Replace(sanitisedNamePart, reservedWordPattern, "_reservedWord_.", RegexOptions.IgnoreCase);
            }

            return sanitisedNamePart;
        }

        public static List<string> GetValidFilePaths(string basePath, IList<string> extensions)
        {
            var results = new List<string>();

            foreach (var path in Directory.GetFiles(basePath))
            {
                if (Flow.Interrupted)
                    break;

                var extension = Path.GetExtension(path);

                if (extensions.Contains(extension))
                {
                    var attributes = File.GetAttributes(path);

                    if (!(attributes.HasFlag(FileAttributes.ReadOnly)
                        || attributes.HasFlag(FileAttributes.Hidden)))
                    {
                        results.Add(path);
                    }
                }
            }

            return results;
        }
    }
}
